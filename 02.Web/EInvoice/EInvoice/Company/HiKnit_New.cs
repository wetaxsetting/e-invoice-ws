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
    public class HiKnit_New
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
                lb_amount_trans = dt.Rows[0]["EXCHANGERATE_NO"].ToString();
                amount_trans = dt.Rows[0]["TOT_AMT_BK_93"].ToString();
                amount_total = dt.Rows[0]["tot_amt_tr_94"].ToString();
                amount_vat = dt.Rows[0]["VAT_TR_AMT_DIS_TR_91"].ToString();
                amount_net = dt.Rows[0]["NET_TR_AMT_DIS_TR_89"].ToString();

                // read_prive = Num2VNText(dt.Rows[0]["TotalAmountInWord"].ToString(), "USD");
            }
            read_prive = dt.Rows[0]["AMOUNT_WORD_VIE"].ToString();//read_prive.Replace(",", "").Replace("TRừ", "Trừ");

        
            string[] parts = dt.Rows[0]["Seller_Address"].ToString().Split(',');
            string l_add = parts[0] + "," + parts[1] + "," + parts[2] + ",";
            string l_add1 = parts[3] + "," + parts[4];
            //read_en = dt.Rows[0]["TotalAmountInWord"].ToString();
            int end = 0;
            int count = count_page_v + r;
            double height = 130;
            StringBuilder htmlStr = new StringBuilder("");
            string heigh = "", heigh_d = "";


            htmlStr.Append("	<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>									 \n");
            htmlStr.Append("	<html>                                                                                                                                   \n");
            htmlStr.Append("	<head>                                                                                                                                   \n");
            htmlStr.Append("	<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>                                                                      \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	<script type='text/javascript'                                                                                                           \n");
            htmlStr.Append("		src='${pageContext.request.contextPath}/system/syscommand.js'></script>                                                              \n");
            htmlStr.Append("	<title>Report E-Invoice</title>                                                                                                          \n");
            htmlStr.Append("	<!-- Set page size here: A5, A4 or A3 -->                                                                                                \n");
            htmlStr.Append("	<!-- Set also 'landscape' if you need -->                                                                                                \n");
            htmlStr.Append("	<style>                                                                                                                                  \n");
            htmlStr.Append("	@page {                                                                                                                                  \n");
            htmlStr.Append("		size: A4                                                                                                                             \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	</style>                                                                                                                                 \n");
            //htmlStr.Append("	<link href='https://fonts.googleapis.com/css?family=Tangerine:700'                                                                       \n");
            //htmlStr.Append("		rel='stylesheet' type='text/css'>                                                                                                    \n");
            htmlStr.Append("	<style>                                                                                                                                  \n");
            htmlStr.Append("	/*body   { font-family: serif }                                                                                                          \n");
            htmlStr.Append("	    h1     { font-family: 'Tangerine', cursive; font-size: 40pt; line-height: 18mm}                                                      \n");
            htmlStr.Append("	    h2, h3 { font-family: 'Tangerine', cursive; font-size: 24pt; line-height: 7mm }                                                      \n");
            htmlStr.Append("	    h4     { font-size: 13pt; line-height: 1mm }                                                                                         \n");
            htmlStr.Append("	    h2 + p { font-size: 18pt; line-height: 7mm }                                                                                         \n");
            htmlStr.Append("	    h3 + p { font-size: 14pt; line-height: 7mm }                                                                                         \n");
            htmlStr.Append("	    li     { font-size: 11pt; line-height: 5mm }                                                                                         \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	    h1      { margin: 0 }                                                                                                                \n");
            htmlStr.Append("	    h1 + ul { margin: 2mm 0 5mm }                                                                                                        \n");
            htmlStr.Append("	    h2, h3  { margin: 0 3mm 3mm 0; float: left }                                                                                         \n");
            htmlStr.Append("	    h2 + p,                                                                                                                              \n");
            htmlStr.Append("	    h3 + p  { margin: 0 0 3mm 50mm }                                                                                                     \n");
            htmlStr.Append("	    //h4      { margin: 1mm 0 0 2mm; border-bottom: 1px solid black }                                                                    \n");
            htmlStr.Append("	    h4 + ul { margin: 5mm 0 0 50mm }                                                                                                     \n");
            htmlStr.Append("	    article { border: 4px double black; padding: 5mm 10mm; border-radius: 3mm }*/                                                        \n");
            htmlStr.Append("	body {                                                                                                                                   \n");
            htmlStr.Append("		color: blue;                                                                                                                         \n");
            htmlStr.Append("		font-size: 100%;                                                                                                                     \n");
            htmlStr.Append("		background-image: url('assets/Solution.jpg');                                                                                        \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	h1 {                                                                                                                                     \n");
            htmlStr.Append("		color: #00FF00;                                                                                                                      \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	p {                                                                                                                                      \n");
            htmlStr.Append("		color: rgb(0, 0, 255)                                                                                                                \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	headline1 {                                                                                                                              \n");
            htmlStr.Append("		background-image: url(assets/Solution.jpg);                                                                                          \n");
            htmlStr.Append("		background-repeat: no-repeat;                                                                                                        \n");
            htmlStr.Append("		background-position: left top;                                                                                                       \n");
            htmlStr.Append("		padding-top: 68px;                                                                                                                   \n");
            htmlStr.Append("		margin-bottom: 50px;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	headline2 {                                                                                                                              \n");
            htmlStr.Append("		background-image: url(images/newsletter_headline2.gif);                                                                              \n");
            htmlStr.Append("		background-repeat: no-repeat;                                                                                                        \n");
            htmlStr.Append("		background-position: left top;                                                                                                       \n");
            htmlStr.Append("		padding-top: 68px;                                                                                                                   \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	<!--                                                                                                                                     \n");
            htmlStr.Append("	[if !mso]> <style>:* {                                                                                                                 \n");
            htmlStr.Append("		behavior: url(#default#VML);                                                                                                         \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	o\\:* {                                                                                                                                   \n");
            htmlStr.Append("		behavior: url(#default#VML);                                                                                                         \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	x\\:* {                                                                                                                                   \n");
            htmlStr.Append("		behavior: url(#default#VML);                                                                                                         \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.shape {                                                                                                                                 \n");
            htmlStr.Append("		behavior: url(#default#VML);                                                                                                         \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	</style>                                                                                                                                 \n");
            htmlStr.Append("	<![endif]-->                                                                                                                             \n");
            htmlStr.Append("	<style id='E invoice - form Hi Knit_27648_Styles'>                                                                                       \n");
            htmlStr.Append("	<!--                                                                                                                                     \n");
            htmlStr.Append("	table {                                                                                                                                  \n");
            htmlStr.Append("		mso-displayed-decimal-separator: '\\.';                                                                                               \n");
            htmlStr.Append("		mso-displayed-thousand-separator: '\\,';                                                                                              \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font522200 {                                                                                                                            \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font622200 {                                                                                                                            \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font722200 {                                                                                                                            \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font822200 {                                                                                                                            \n");
            htmlStr.Append("		color: black;                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font922200 {                                                                                                                            \n");
            htmlStr.Append("		color: black;                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font1022200 {                                                                                                                           \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 8pt;                                                                                                                      \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font1122200 {                                                                                                                           \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font1222200 {                                                                                                                           \n");
            htmlStr.Append("		color: #0066CC;                                                                                                                      \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font1322200 {                                                                                                                           \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font1422200 {                                                                                                                           \n");
            htmlStr.Append("		color: #16365C;                                                                                                                      \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font1522200 {                                                                                                                           \n");
            htmlStr.Append("		color: #16365C;                                                                                                                      \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font1622200 {                                                                                                                           \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 9.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: Cambria, serif;                                                                                                         \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font1722200 {                                                                                                                           \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 9.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: Century, serif;                                                                                                         \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font1822200 {                                                                                                                           \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                 \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font1922200 {                                                                                                                           \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                 \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.font2022200 {                                                                                                                           \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 11.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl6322200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl6422200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl6522200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 14.5pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: top;                                                                                                                 \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl6622200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 14.5pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl6722200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 14.5pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl6822200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                        \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl6922200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: navy;                                                                                                                         \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl7022200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                        \n");
            htmlStr.Append("		font-size: 0.9pt;                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 2.0pt double windowtext;                                                                                                 \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl7122200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                        \n");
            htmlStr.Append("		font-size: 0.9pt;                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 2.0pt double windowtext;                                                                                                 \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl7222200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: navy;                                                                                                                         \n");
            htmlStr.Append("		font-size: 0.9pt;                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 2.0pt double windowtext;                                                                                                 \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl7322200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: navy;                                                                                                                         \n");
            htmlStr.Append("		font-size: 0.9pt;                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 2.0pt double windowtext;                                                                                                 \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl7422200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                        \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl7522200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: navy;                                                                                                                         \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl7622200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                        \n");
            htmlStr.Append("		font-size: 0.9pt;                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl7722200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                        \n");
            htmlStr.Append("		font-size: 0.9pt;                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl7822200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: navy;                                                                                                                         \n");
            htmlStr.Append("		font-size: 2.0pt;                                                                                                                    \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl7922200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: navy;                                                                                                                         \n");
            htmlStr.Append("		font-size: 0.9pt;                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl8022200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: navy;                                                                                                                         \n");
            htmlStr.Append("		font-size: 0.9pt;                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl8122200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl8222200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 14.5pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: top;                                                                                                                 \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl8322200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 14.5pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: top;                                                                                                                 \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl8422200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: top;                                                                                                                 \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl8522200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl8622200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl8722200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl8822200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl8922200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl9022200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl9122200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl9222200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl9322200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl9422200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl9522200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl9622200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl9722200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl9822200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl9922200 {                                                                                                                             \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl10022200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl10122200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl10222200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl10322200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: #0070C0;                                                                                                                      \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl10422200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("		mso-text-control: shrinktofit;                                                                                                       \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl10522200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl10622200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl10722200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl10822200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl10922200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl11022200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: black;                                                                                                                        \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl11122200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: navy;                                                                                                                         \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl11222200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: navy;                                                                                                                         \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl11322200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: #C00000;                                                                                                                      \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl11422200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: top;                                                                                                                 \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl11522200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: top;                                                                                                                 \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl11622200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl11722200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl11822200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl11922200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl12022200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl12122200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl12222200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                 \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl12322200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                 \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl12422200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                 \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: general;                                                                                                                 \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl12522200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl12622200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl12722200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl12822200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl12922200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl13022200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl13122200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl13222200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl13322200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl133222002 {                                                                                                                           \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		border-top: 0.5 dotted windowtext;                                                                                                   \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl13422200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl13522200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl13622200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl13722200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl13822200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("		mso-text-control: shrinktofit;                                                                                                       \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl13922200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("		mso-text-control: shrinktofit;                                                                                                       \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl14022200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: red;                                                                                                                          \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("		mso-text-control: shrinktofit;                                                                                                       \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl14122200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 9.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: Cambria, serif;                                                                                                         \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 2.0pt double windowtext;                                                                                              \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl14222200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 9.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: Cambria, serif;                                                                                                         \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 2.0pt double windowtext;                                                                                              \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl14322200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 9.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: Cambria, serif;                                                                                                         \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 2.0pt double windowtext;                                                                                              \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl14422200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl14522200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl14622200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl14722200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl14822200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl14922200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl15022200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl15122200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl15222200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl15322200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl15422200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl15522200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl15622200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl15722200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl15822200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl15922200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl16022200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl16122200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border: 1pt solid windowtext;                                                                                                       \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl16222200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border: 1pt solid windowtext;                                                                                                       \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl16322200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl16422200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl16522200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl16622200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl16722200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: right;                                                                                                                   \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl16822200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl16922200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border: 1pt solid windowtext;                                                                                                       \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl17022200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl17122200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl17222200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl17322200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 20.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl17422200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 20.0pt;                                                                                                                  \n");
            htmlStr.Append("		font-weight: 700;                                                                                                                    \n");
            htmlStr.Append("		font-style: italic;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl17522200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl17622200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl17722200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl17822200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl17922200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.5pt;                                                                                                                 \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: top;                                                                                                                 \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl18022200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl18122200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: left;                                                                                                                    \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: normal;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl18222200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt solid windowtext;                                                                                                   \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: none;                                                                                                                 \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl18322200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: 1pt solid windowtext;                                                                                                  \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl18422200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl18522200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: middle;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: 1pt solid windowtext;                                                                                                 \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl18622200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 1pt dotted windowtext;                                                                                               \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl18722200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 0.9pt;                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl18822200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl18922200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 14.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: none;                                                                                                                    \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl19022200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 0.2pt;                                                                                                                    \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: 2.0pt double windowtext;                                                                                                \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl19122200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: none;                                                                                                                  \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	.xl19222200 {                                                                                                                            \n");
            htmlStr.Append("		padding: 0px;                                                                                                                        \n");
            htmlStr.Append("		mso-ignore: padding;                                                                                                                 \n");
            htmlStr.Append("		color: windowtext;                                                                                                                   \n");
            htmlStr.Append("		font-size: 12.0pt;                                                                                                                   \n");
            htmlStr.Append("		font-weight: 400;                                                                                                                    \n");
            htmlStr.Append("		font-style: normal;                                                                                                                  \n");
            htmlStr.Append("		text-decoration: none;                                                                                                               \n");
            htmlStr.Append("		font-family: 'Times New Roman', serif;                                                                                               \n");
            htmlStr.Append("		mso-font-charset: 0;                                                                                                                 \n");
            htmlStr.Append("		mso-number-format: General;                                                                                                          \n");
            htmlStr.Append("		text-align: center;                                                                                                                  \n");
            htmlStr.Append("		vertical-align: bottom;                                                                                                              \n");
            htmlStr.Append("		border-top: 1pt dotted windowtext;                                                                                                  \n");
            htmlStr.Append("		border-right: 2.0pt double windowtext;                                                                                               \n");
            htmlStr.Append("		border-bottom: 1pt solid windowtext;                                                                                                \n");
            htmlStr.Append("		border-left: none;                                                                                                                   \n");
            htmlStr.Append("		background: white;                                                                                                                   \n");
            htmlStr.Append("		mso-pattern: black none;                                                                                                             \n");
            htmlStr.Append("		white-space: nowrap;                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                        \n");
            htmlStr.Append("	-->                                                                                                                                      \n");
            htmlStr.Append("	</style>                                                                                                                                 \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("	</head>                                                                                                                                  \n");
            htmlStr.Append("	<body class='A4'>                                                                                                                        \n");

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

            double v_totalHeightLastPage = 203.5;// 243.5

            double v_totalHeightPage = 580;//   540;

            for (int k = 0; k < v_countNumberOfPages; k++)
            {
                v_totalHeightPage = 450;// 540;

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

                htmlStr.Append("		<div id='E invoice - form Hi Knit_22200' align=center                                                                                \n");
                htmlStr.Append("			x:publishsource='Excel'>                                                                                                         \n");
                htmlStr.Append("			<table>                                                                                                                      \n");
                htmlStr.Append("				<tr class= height=6                                                                                                 \n");
                htmlStr.Append("					style='mso-height-source: userset; height: 10.5pt'>                                                                       \n");
                htmlStr.Append("					<td height=6 class= style='height: 10.5pt'>&nbsp;</td>                                                           \n");
                htmlStr.Append("				    </tr>                                                                                                                        \n");
                htmlStr.Append("			</table>                                                                                                                      \n");
                htmlStr.Append("			<table border=0 cellpadding=0 cellspacing=0 width=821 class=xl6322200                                                            \n");
                htmlStr.Append("			 style='border-collapse:collapse;table-layout:fixed;width:622pt'>                                                                \n");
                htmlStr.Append("			 <col class=xl6322200 width=59 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 2104;width:44pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=107 style='mso-width-source:userset;mso-width-alt:                                                   \n");
                htmlStr.Append("			 3811;width:80pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=23 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 824;width:17pt'>                                                                                                                \n");
                htmlStr.Append("			 <col class=xl6322200 width=46 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 1621;width:34pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=40 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 1422;width:30pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=29 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 1024;width:22pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=34 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 1223;width:26pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=25 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 881;width:19pt'>                                                                                                                \n");
                htmlStr.Append("			 <col class=xl6322200 width=120 style='mso-width-source:userset;mso-width-alt:                                                   \n");
                htmlStr.Append("			 4266;width:90pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=19 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 682;width:14pt'>                                                                                                                \n");
                htmlStr.Append("			 <col class=xl6322200 width=30 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 1080;width:23pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=41 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 1450;width:31pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=38 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 1365;width:39pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=49 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 1735;width:27pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=38 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 1336;width:28pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=106 style='mso-width-source:userset;mso-width-alt:                                                   \n");
                htmlStr.Append("			 3754;width:79pt'>                                                                                                               \n");
                htmlStr.Append("			 <col class=xl6322200 width=17 style='mso-width-source:userset;mso-width-alt:                                                    \n");
                htmlStr.Append("			 597;width:19pt'>                                                                                                                \n");
                htmlStr.Append("	                                                                                                                                         \n");
                htmlStr.Append("				<tr class=xl6822200 height=6                                                                                                 \n");
                htmlStr.Append("					style='mso-height-source: userset; height: 0.5pt'>                                                                       \n");
                htmlStr.Append("					<td height=6 class=xl7022200 style='height: 0.5pt'>&nbsp;</td>                                                           \n");
                htmlStr.Append("					<td class=xl7122200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7122200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7122200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7222200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7222200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7222200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7222200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7222200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7222200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7222200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7222200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7222200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7222200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7222200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7222200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7322200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr class=xl6822200 height=22                                                                                                \n");
                htmlStr.Append("					style='mso-height-source: userset; height: 21.0pt'>                                                                     \n");
                htmlStr.Append("					<td height=22 style='height: 22.0pt;border-left: 2.0pt double windowtext;' align=left valign=top><![if !vml]><span      \n");
                htmlStr.Append("						style='mso-ignore: vglayout; position: absolute; z-index: 1; margin-left: 10px; margin-top: 2px; width: 165px; height: 81px'><img \n");
                htmlStr.Append("							width=165 height=81                                                                                              \n");
                htmlStr.Append("							src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\HiKnit.png'                                                       \n");
                htmlStr.Append("							v:shapes='Picture_x0020_1'></span>                                                                               \n");
                htmlStr.Append("					<![endif]><span style='mso-ignore: vglayout2'>                                                                           \n");
                htmlStr.Append("							<table cellpadding=0 cellspacing=0>                                                                              \n");
                htmlStr.Append("								<tr>                                                                                                         \n");
                htmlStr.Append("									<td height=22 class=xl7422200 width=59                                                                   \n");
                htmlStr.Append("										style='height: 22.0pt; width: 37.4pt;border-left: none;'>&nbsp;</td>                                \n");
                htmlStr.Append("								</tr>                                                                                                        \n");
                htmlStr.Append("							</table>                                                                                                         \n");
                htmlStr.Append("					</span></td>                                                                                                             \n");
                htmlStr.Append("					<td class=xl6822200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl6822200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl8622200 colspan=11>&#272;&#417;n v&#7883; bán hàng                                                           \n");
                htmlStr.Append("						(<font class='font622200'>Sale company</font><font                                                                   \n");
                htmlStr.Append("						class='font522200'>): </font><font class='font1522200'>" + dt.Rows[0]["Seller_Name"] + "<span                                        \n");
                htmlStr.Append("							style='mso-spacerun: yes'> </span></font>                                                                        \n");
                htmlStr.Append("					</td>                                                                                                                    \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl7522200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr class=xl6822200 height=22                                                                                                \n");
                htmlStr.Append("					style='mso-height-source: userset; height: 21.0pt'>                                                                     \n");
                htmlStr.Append("					<td height=22 class=xl7422200 style='height: 21.0pt'>&nbsp;</td>                                                        \n");
                htmlStr.Append("					<td class=xl6822200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl6822200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl8622200 colspan=6>Mã S&#7889; Thu&#7871; (<font                                                              \n");
                htmlStr.Append("						class='font622200'>Tax code</font><font class='font522200'>):                                                        \n");
                htmlStr.Append("					</font><font class='font1422200'>" + dt.Rows[0]["Seller_TaxCode"] + "</font></td>                                                      \n");
                htmlStr.Append("					<td class=xl11022200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11022200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl7522200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr class=xl6822200 height=22                                                                                                \n");
                htmlStr.Append("					style='mso-height-source: userset; height: 21.0pt'>                                                                     \n");
                htmlStr.Append("					<td height=22 class=xl7422200 style='height: 21.0pt'>&nbsp;</td>                                                        \n");
                htmlStr.Append("					<td class=xl6822200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl6822200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl8622200 colspan=13>&#272;&#7883;a Ch&#7881; (<font                                                           \n");
                htmlStr.Append("						class='font622200'>Address</font><font class='font522200'>):                                                         \n");
                htmlStr.Append("							" + l_add + "<span style='mso-spacerun: yes'> </span>                                                      \n");
                htmlStr.Append("					</font></td>                                                                                                             \n");
                htmlStr.Append("					<td class=xl7522200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr class=xl6822200 height=22                                                                                                \n");
                htmlStr.Append("					style='mso-height-source: userset; height: 21.0pt'>                                                                     \n");
                htmlStr.Append("					<td height=22 class=xl7422200 style='height: 21.0pt'>&nbsp;</td>                                                        \n");
                htmlStr.Append("					<td class=xl6822200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl6822200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl8622200 colspan=9><span style='mso-spacerun: yes'>                                                           \n");
                htmlStr.Append("					</span>" + l_add1 + "</td>                                                                                           \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl7522200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr class=xl6822200 height=22                                                                                                \n");
                htmlStr.Append("					style='mso-height-source: userset; height: 21.0pt'>                                                                     \n");
                htmlStr.Append("					<td height=22 class=xl7422200 style='height: 21.0pt'>&nbsp;</td>                                                        \n");
                htmlStr.Append("					<td class=xl6822200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl6822200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl8622200 colspan=6><span style='mso-spacerun: yes'> </span>&#272;i&#7879;n                                    \n");
                htmlStr.Append("						Tho&#7841;i (<font class='font622200'>Tel</font><font                                                                \n");
                htmlStr.Append("						class='font522200'>): </font><font class='font922200'>" + dt.Rows[0]["Seller_Tel"] + ";<span                                       \n");
                htmlStr.Append("							style='mso-spacerun: yes'>  </span></font></td>                                                                  \n");
                htmlStr.Append("					<td class=xl11022200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11022200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11122200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl7522200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("    		  <tr class=xl6822200 height=22 style='mso-height-source:userset;height:16.95pt'>                                                               \n");
                htmlStr.Append("      <td height=22 class=xl7422200 style='height:16.95pt'>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl6822200>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl6822200>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl8622200 colspan=13><span                                                               \n");
                htmlStr.Append("      style='mso-spacerun:yes'> </span><font                                                               \n");
                htmlStr.Append("      class='font922200'>" + dt.Rows[0]["Seller_Accountno"] + " " + dt.Rows[0]["BANK_NM78"] + " <span style='mso-spacerun:yes'>  </span></font></td>                                                               \n");
                //htmlStr.Append("      <td class=xl11022200>&nbsp;</td>                                                               \n");
                //htmlStr.Append("      <td class=xl11022200>&nbsp;</td>                                                               \n");
                //htmlStr.Append("      <td class=xl11122200>&nbsp;</td>                                                               \n");
                //htmlStr.Append("      <td class=xl11122200>&nbsp;</td>                                                               \n");
                //htmlStr.Append("      <td class=xl11122200>&nbsp;</td>                                                               \n");
                //htmlStr.Append("      <td class=xl11122200>&nbsp;</td>                                                               \n");
                //htmlStr.Append("      <td class=xl11122200>&nbsp;</td>                                                               \n");
                htmlStr.Append("      <td class=xl7522200>&nbsp;</td>                                                               \n");
                htmlStr.Append("     </tr>                                                               \n");
                htmlStr.Append("				<tr class=xl6822200 height=6                                                                                                 \n");
                htmlStr.Append("					style='mso-height-source: userset; height: 0.95pt'>                                                                      \n");
                htmlStr.Append("					<td height=6 class=xl7622200 style='height: 0.95pt'>&nbsp;</td>                                                          \n");
                htmlStr.Append("					<td class=xl7722200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7722200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7722200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7722200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7722200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7722200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7722200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7822200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7722200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7722200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7922200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7922200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7922200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7922200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl7922200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl8022200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr height=40 style='mso-height-source: userset; height: 30.0pt'>                                                            \n");
                htmlStr.Append("					<td height=40 class=xl6422200 style='height: 30.0pt'>&nbsp;</td>                                                         \n");
                htmlStr.Append("					<td class=xl6622200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl6622200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td colspan=10 class=xl17322200>HÓA &#272;&#416;N GIÁ TR&#7882;                                                          \n");
                htmlStr.Append("						GIA T&#258;NG</td>                                                                                                   \n");
                htmlStr.Append("					<td class=xl12222200 colspan=4                                                                                           \n");
                htmlStr.Append("						style='border-right: 2.0pt double black'><font                                                    \n");
                htmlStr.Append("						class='font1922200'></font><font class='font1822200'></font></td>                                \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr height=30 style='height: 22.2pt'>                                                                                       \n");
                htmlStr.Append("					<td height=30 class=xl6522200 style='height: 22.2pt'>&nbsp;</td>                                                        \n");
                htmlStr.Append("					<td class=xl6622200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl6622200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td colspan=10 class=xl17422200><span style='mso-spacerun: yes'> </span>(Value                                           \n");
                htmlStr.Append("						added Invoice)</td>                                                                                                  \n");
                htmlStr.Append("					<td class=xl12222200 colspan=3>Ký hi&#7879;u/<font                                                                       \n");
                htmlStr.Append("						class='font1922200'>Serial</font><font class='font1822200'>:                                                         \n");
                htmlStr.Append("							" + dt.Rows[0]["templateCode"] + "" + dt.Rows[0]["InvoiceSerialNo"] + "</font></td>                                                                                     \n");
                htmlStr.Append("					<td class=xl9422200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr height=24 style='mso-height-source: userset; height: 21.0pt'>                                                           \n");
                htmlStr.Append("					<td height=24 class=xl6722200 style='height: 21.0pt'>&nbsp;</td>                                                        \n");
                htmlStr.Append("					<td class=xl6622200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl6622200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td colspan=10 class=xl13422200></td>                                                                                    \n");
                htmlStr.Append("					<td class=xl12322200 colspan=2>S&#7889;/<font                                                                            \n");
                htmlStr.Append("						class='font1922200'>No</font><font class='font1822200'>:</font></td>                                                 \n");
                htmlStr.Append("					<td class=xl12422200>" + dt.Rows[0]["InvoiceNumber"] + "</td>                                                                             \n");
                htmlStr.Append("					<td class=xl9422200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("	                                                                                                                                         \n");
                htmlStr.Append("				<tr class=xl8422200 height=25 style='height: 18.6pt'>                                                                        \n");
                htmlStr.Append("					<td height=25 class=xl8222200 style='height: 18.6pt'>&nbsp;</td>                                                         \n");
                htmlStr.Append("					<td class=xl8322200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td class=xl8322200>&nbsp;</td>                                                                                          \n");
                htmlStr.Append("					<td colspan=10 class=xl17922200>Ngày (<font class='font1922200'>date</font><font                                         \n");
                htmlStr.Append("						class='font1822200'>)<span style='mso-spacerun: yes'>&nbsp;" + dt.Rows[0]["invoiceissueddate_dd"] + "&nbsp;&nbsp;&nbsp;                          \n");
                htmlStr.Append("						</span>tháng (                                                                                                       \n");
                htmlStr.Append("					</font><font class='font1922200'>month</font><font class='font1822200'>)<span                                            \n");
                htmlStr.Append("							style='mso-spacerun: yes'>&nbsp;" + dt.Rows[0]["invoiceissueddate_mm"] + "&nbsp;&nbsp;&nbsp;                                                 \n");
                htmlStr.Append("						</span>N&#259;m (                                                                                                    \n");
                htmlStr.Append("					</font><font class='font1922200'> year</font><font class='font1822200'>)<span                                            \n");
                htmlStr.Append("							style='mso-spacerun: yes'>&nbsp;" + dt.Rows[0]["invoiceissueddate_yyyy"] + "&nbsp;&nbsp;&nbsp;                                               \n");
                htmlStr.Append("						</span></font></td>                                                                                                  \n");
                htmlStr.Append("					<td class=xl11422200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11422200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11422200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("					<td class=xl11522200>&nbsp;</td>                                                                                         \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr height=22 style='mso-height-source: userset; height: 21.0pt'>                                                           \n");
                htmlStr.Append("					<td colspan=4 height=22 class=xl17522200 style='height: 21.0pt'><span                                                   \n");
                htmlStr.Append("						style='mso-spacerun: yes'>  </span>H&#7885; tên ng&#432;&#7901;i                                                     \n");
                htmlStr.Append("						mua<font class='font622200'>(Buyer's name)</font><font                                                               \n");
                htmlStr.Append("						class='font522200'>:</font></td>                                                                                     \n");
                htmlStr.Append("					<td colspan=13 class=xl17722200                                                                                          \n");
                htmlStr.Append("						style='border-right: 2.0pt double black'>&nbsp;" + dt.Rows[0]["buyer"] + "</td>                                              \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr height=22 style='mso-height-source: userset; height: 21.0pt'>                                                           \n");
                htmlStr.Append("					<td height=22 class=xl9722200 colspan=2 style='height: 21.0pt'><span                                                    \n");
                htmlStr.Append("						style='mso-spacerun: yes'>  </span>&#272;&#7883;a ch&#7881; <font                                                    \n");
                htmlStr.Append("						class='font622200'>(Address)</font><font class='font522200'>:</font></td>                                            \n");
                htmlStr.Append("					<td colspan=15 class=xl17222200                                                                                          \n");
                htmlStr.Append("						style='border-right: 2.0pt double black'>" + dt.Rows[0]["Attribute_05"] + "</td>                                                      \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr height=22 style='mso-height-source: userset; height: 21.0pt'>                                                           \n");
                htmlStr.Append("					<td height=22 class=xl9722200 colspan=2 style='height: 21.0pt'><span                                                    \n");
                htmlStr.Append("						style='mso-spacerun: yes'>  </span>Tên &#273;&#417;n v&#7883; (<font                                                 \n");
                htmlStr.Append("						class='font2022200'>Company</font><font class='font522200'>)                                                         \n");
                htmlStr.Append("							:</font></td>                                                                                                    \n");
                htmlStr.Append("					<td colspan=15 class=xl18122200                                                                                          \n");
                htmlStr.Append("						style='border-right: 2.0pt double black'>&nbsp;" + dt.Rows[0]["buyerlegalname"] + "</td>                                               \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr height=22 style='mso-height-source: userset; height: 21.0pt'>                                                           \n");
                htmlStr.Append("					<td height=22 class=xl9722200 colspan=2 style='height: 21.0pt'><span                                                    \n");
                htmlStr.Append("						style='mso-spacerun: yes'>  </span>&#272;&#7883;a ch&#7881; <font                                                    \n");
                htmlStr.Append("						class='font622200'>(Address)</font><font class='font522200'>:</font></td>                                            \n");
                htmlStr.Append("					<td colspan=15 class=xl18122200                                                                                          \n");
                htmlStr.Append("						style='border-right: 2.0pt double black'>&nbsp;" + dt.Rows[0]["BuyerAddress"] + "</td>                                                  \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("		<tr height=22 style='mso-height-source:userset;height:16.95pt'>																 \n");
                htmlStr.Append("			<td height=22 class=xl9722200 colspan=2 style='height:16.95pt'><span                                                     \n");
                htmlStr.Append("			style='mso-spacerun:yes'>  </span>MST<font class='font622200'>(Tax code)</font><font                                     \n");
                htmlStr.Append("			class='font522200'>: &nbsp;</font></td>                                                                                  \n");
                htmlStr.Append("			<td colspan=6 class=xl18122200 >" + dt.Rows[0]["BuyerTaxCode"] + "&nbsp;</td>                                                              \n");
                htmlStr.Append("			<td colspan=6 class=xl18122200 style='border-bottom:none'><span                                                          \n");
                htmlStr.Append("			style='mso-spacerun:yes'>  </span>Hình th&#7913;c thanh toán (<font                                                      \n");
                htmlStr.Append("			class='font622200'>Method of payment</font><font class='font522200'>) :<span                                             \n");
                htmlStr.Append("			style='mso-spacerun:yes'></span></font></td>                                                                             \n");
                htmlStr.Append("			<td colspan=3 class=xl18122200 style='border-right:2.0pt double black'>&nbsp;" + dt.Rows[0]["PaymentMethodCK"] + "</td>                 \n");
                htmlStr.Append("		</tr>                                                                                                                        \n"); 
                htmlStr.Append("     </tr>                                                               \n");
                /*  htmlStr.Append("				<tr height=22 style='mso-height-source: userset; height: 21.0pt'>                                                           \n");
                  htmlStr.Append("					<td height=22 class=xl9722200 colspan=2 style='height: 21.0pt'><span                                                    \n");
                  htmlStr.Append("						style='mso-spacerun: yes'>  </span>MST<font class='font622200'>(Tax                                                  \n");
                  htmlStr.Append("							code)</font><font class='font522200'>:</font></td>                                                               \n");
                  htmlStr.Append("					<td colspan=15 class=xl18122200                                                                                          \n");
                  htmlStr.Append("						style='border-right: 2.0pt double black'>" + dt.Rows[0]["BuyerTaxCode"] + "&nbsp;</td>                                                 \n");
                  htmlStr.Append("				</tr>                                                                                                                        \n");
                  htmlStr.Append("				<tr class=xl6322200 height=22                                                                                                \n");
                  htmlStr.Append("					style='mso-height-source: userset; height: 21.0pt'>                                                                     \n");
                  htmlStr.Append("					<td height=22 class=xl8522200 colspan=6 style='height: 22.0pt;'><span                                                    \n");
                  htmlStr.Append("						style='mso-spacerun: yes'>  </span>Hình th&#7913;c thanh toán (<font                                                 \n");
                  htmlStr.Append("						class='font622200'>Method of payment</font><font class='font522200'>)                                                \n");
                  htmlStr.Append("							:</td>                                                                                                           \n");
                  htmlStr.Append("					<td colspan=2 class=xl17222200 style='border-right: 2.0pt double white'>&nbsp;" + dt.Rows[0]["PaymentMethodCK"] + "</td>                                                         \n");
                  htmlStr.Append("					<td colspan=4 class=xl13322200                                                                                           \n");
                  htmlStr.Append("						>S&#7889; tài                                                                \n");
                  htmlStr.Append("						kho&#7843;n (<font class='font2022200'>Account No</font><font                                                        \n");
                  htmlStr.Append("						class='font522200'>.) :</font>                                                                                       \n");
                  htmlStr.Append("					</td>                                                                                                                    \n");
                  htmlStr.Append("					<td colspan=5 class=xl18122200                                                                                           \n");
                  htmlStr.Append("						style='border-right: 2.0pt double black'>&nbsp;</td>                                                                 \n");
                  htmlStr.Append("				</tr>                                                                                                                        \n");*/
                htmlStr.Append("				<tr height=22 style='mso-height-source: userset; height: 21.0pt'>                                                           \n");
                htmlStr.Append("					<td height=22 class=xl9722200 colspan=4 style='height: 21.0pt'><span                                                    \n");
                htmlStr.Append("						style='mso-spacerun: yes'>  </span>Giao hàng t&#7841;i kho (<font                                                    \n");
                htmlStr.Append("						class='font2022200'>Delivery place</font><font class='font522200'>)                                                  \n");
                htmlStr.Append("					</font><span style='display: none'><font class='font522200'>:</font></span></td>                                         \n");
                htmlStr.Append("					<td colspan=13 class=xl17222200                                                                                          \n");
                htmlStr.Append("						style='border-right: 2.0pt double black;border-top:none'>&nbsp;" + dt.Rows[0]["Attribute_01"] + "</td>                                         \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr height=22 style='mso-height-source: userset; height: 21.0pt'>                                                           \n");
                htmlStr.Append("					<td height=22 class=xl9722200 colspan=2 style='height: 21.0pt'><span                                                    \n");
                htmlStr.Append("						style='mso-spacerun: yes'>  </span>S&#7889; order <font                                                              \n");
                htmlStr.Append("						class='font622200'>(No)</font><font class='font522200'>:</font></td>                                                 \n");
                htmlStr.Append("					<td colspan=15 class=xl17222200                                                                                          \n");
                htmlStr.Append("						style='border-right: 2.0pt double black'>&nbsp;" + dt.Rows[0]["Attribute_04"] + "</td>                                                       \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr height=6 style='mso-height-source: userset; height: 2.95pt'>                                                             \n");
                htmlStr.Append("					<td colspan=17 height=6 class=xl18722200                                                                                 \n");
                htmlStr.Append("						style='border-right: 2.0pt double black; height: 2.95pt'>&nbsp;</td>                                                 \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr class=xl8122200 height=49                                                                                                \n");
                htmlStr.Append("					style='mso-height-source: userset; height: 36.75pt'>                                                                     \n");
                htmlStr.Append("					<td height=49 class=xl8722200 width=59                                                                                   \n");
                htmlStr.Append("						style='height: 36.75pt; border-top: none; width: 37.4pt'>STT<br>                                                     \n");
                htmlStr.Append("						<font class='font522200'>No</font></td>                                                                              \n");
                htmlStr.Append("					<td colspan=8 class=xl16922200 width=424                                                                                 \n");
                htmlStr.Append("						style='border-left: none; width: 318pt'>Tên hàng hoá,                                                                \n");
                htmlStr.Append("						d&#7883;ch v&#7909;<br> <font class='font622200'>Commodity</font>                                                    \n");
                htmlStr.Append("					</td>                                                                                                                    \n");
                htmlStr.Append("					<td colspan=2 class=xl16922200 width=49                                                                                  \n");
                htmlStr.Append("						style='border-left: none; width: 37pt'>&#272;VT<br> <font                                                            \n");
                htmlStr.Append("						class='font622200'>Unit</font></td>                                                                                  \n");
                htmlStr.Append("					<td colspan=2 class=xl16922200 width=79                                                                                  \n");
                htmlStr.Append("						style='border-left: none; width: 60pt'>S&#7889;                                                                      \n");
                htmlStr.Append("						l&#432;&#7907;ng<br> <font class='font622200'>Quantities</font>                                                      \n");
                htmlStr.Append("					</td>                                                                                                                    \n");
                htmlStr.Append("					<td colspan=2 class=xl16922200 width=87                                                                                  \n");
                htmlStr.Append("						style='border-left: none; width: 65pt'>&#272;&#417;n giá<br>                                                         \n");
                htmlStr.Append("						<font class='font622200'>Unit price</font></td>                                                                      \n");
                htmlStr.Append("					<td colspan=2 class=xl17022200 width=123                                                                                 \n");
                htmlStr.Append("						style='border-right: 2.0pt double black; border-left: none; width: 92pt'>Thành                                       \n");
                htmlStr.Append("						ti&#7873;n<br> <font class='font622200'>Amount</font>                                                                \n");
                htmlStr.Append("					</td>                                                                                                                    \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");
                htmlStr.Append("				<tr height=26 style='mso-height-source: userset; height: 19.5pt'>                                                            \n");
                htmlStr.Append("					<td height=26 class=xl8822200                                                                                            \n");
                htmlStr.Append("						style='height: 19.5pt; border-top: none'>1</td>                                                                      \n");
                htmlStr.Append("					<td colspan=8 class=xl16122200 width=424                                                                                 \n");
                htmlStr.Append("						style='border-left: none; width: 318pt'>2</td>                                                                       \n");
                htmlStr.Append("					<td colspan=2 class=xl16222200 style='border-left: none'>3</td>                                                          \n");
                htmlStr.Append("					<td colspan=2 class=xl16222200 style='border-left: none'>4</td>                                                          \n");
                htmlStr.Append("					<td colspan=2 class=xl16222200 style='border-left: none'>5</td>                                                          \n");
                htmlStr.Append("					<td colspan=2 class=xl12722200                                                                                           \n");
                htmlStr.Append("						style='border-right: 2.0pt double black; border-left: none;text-align:center;'>6                                     \n");
                htmlStr.Append("						= 4 x 5</td>                                                                                                         \n");
                htmlStr.Append("				</tr>                                                                                                                        \n");


                v_rowHeight = "25.0pt"; //"26.5pt";
                v_rowHeightEmpty = "22.0pt";
                v_rowHeightNumber = 20;

                v_rowHeightLast = "21.0pt";// "23.5pt";
                v_rowHeightLastNumber = 26;// 23.5;
                v_rowHeightEmptyLast = "23.5pt"; //"23.5pt";


                for (int dtR = 0; dtR < page[k]; dtR++)
                {
                    if (!vlongItemName && dt_d.Rows[v_index]["ITEM_NAME"].ToString().Length >= 92)
                    {
                        v_rowHeight = "25.0pt"; //"26.5pt";    
                        v_rowHeightLast = "21.0pt"; //"27.5pt";
                        v_rowHeightLastNumber = 20;//27.5;
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
                        htmlStr.Append("				<tr height=24 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                           \n");
                        htmlStr.Append("					<td height=24 class=xl8922200 style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                           \n");
                        htmlStr.Append("					<td colspan=8 class=xl14522200 width=424                                                                                 \n");
                        htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 318pt;'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                    \n");
                        htmlStr.Append("					<td colspan=2 class=xl16322200 width=49                                                                                  \n");
                        htmlStr.Append("						style='border-left: none; width: 37pt'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                                                      \n");
                        htmlStr.Append("					<td colspan=2 class=xl15022200 width=79 style='width: 60pt'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                     \n");
                        htmlStr.Append("					<td colspan=2 class=xl16522200 width=87                                                                                  \n");
                        htmlStr.Append("						style='border-right: 1pt solid black; width: 65pt'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                         \n");
                        htmlStr.Append("					<td colspan=2 class=xl16722200 width=123                                                                                 \n");
                        htmlStr.Append("						style='border-right: 2.0pt double black; width: 92pt'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                       \n");
                        htmlStr.Append("				</tr>                                                                                                                        \n");

                    }
                    else if (dtR == page[k] - 1)//dong cuoi moi trang
                    {
                        if (k < v_countNumberOfPages - 1) //trang giua
                        {
                            htmlStr.Append("				<tr height=24 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                           \n");
                            htmlStr.Append("					<td height=24 class=xl8922200                                                                                            \n");
                            htmlStr.Append("						style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; border-bottom:none;'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                   \n");
                            htmlStr.Append("					<td colspan=8 class=xl14522200 width=424                                                                                 \n");
                            htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 318pt;border-bottom:none;'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                     \n");
                            htmlStr.Append("					<td colspan=2 class=xl14822200 width=49                                                                                  \n");
                            htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 37pt;border-bottom:none;'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                      \n");
                            htmlStr.Append("					<td colspan=2 class=xl15022200 width=79                                                                                  \n");
                            htmlStr.Append("						style='border-left: none; width: 60pt;border-bottom:none;'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                      \n");
                            htmlStr.Append("					<td colspan=2 class=xl15522200 width=87                                                                                  \n");
                            htmlStr.Append("						style='border-right: 1pt solid black; width: 65pt;border-bottom:none;'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                         \n");
                            htmlStr.Append("					<td colspan=2 class=xl15722200 width=123                                                                                 \n");
                            htmlStr.Append("						style='border-right: 2.0pt double black; border-left: none; width: 92pt;border-bottom:none;'>&nbsp;" + dt_d.Rows[v_index][4] + "&nbsp;</td>                    \n");
                            htmlStr.Append("				</tr>                                                                                                                        \n");


                        }
                        else // trang cuoi
                        {
                            if (dtR == rowsPerPage - 1) // du 11 dong
                            {
                                htmlStr.Append("				<tr height=24 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                           \n");
                                htmlStr.Append("					<td height=24 class=xl8922200                                                                                            \n");
                                htmlStr.Append("						style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + ";'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                   \n");
                                htmlStr.Append("					<td colspan=8 class=xl14522200 width=424                                                                                 \n");
                                htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 318pt;border-bottom:none;'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                     \n");
                                htmlStr.Append("					<td colspan=2 class=xl14822200 width=49                                                                                  \n");
                                htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 37pt;border-bottom:none;'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                      \n");
                                htmlStr.Append("					<td colspan=2 class=xl15022200 width=79                                                                                  \n");
                                htmlStr.Append("						style='border-left: none; width: 60pt;border-bottom:none;'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                      \n");
                                htmlStr.Append("					<td colspan=2 class=xl15522200 width=87                                                                                  \n");
                                htmlStr.Append("						style='border-right: 1pt solid black; width: 65pt;border-bottom:none;'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                         \n");
                                htmlStr.Append("					<td colspan=2 class=xl15722200 width=123                                                                                 \n");
                                htmlStr.Append("						style='border-right: 2.0pt double black; border-left: none; width: 92pt;border-bottom:none;'>&nbsp;" + dt_d.Rows[v_index][4] + "&nbsp;</td>                    \n");
                                htmlStr.Append("				</tr>                                                                                                                        \n");

                            }
                            else
                            {
                                htmlStr.Append("				<tr height=24 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                           \n");
                                htmlStr.Append("					<td height=24 class=xl8922200                                                                                            \n");
                                htmlStr.Append("						style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "; '>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                   \n");
                                htmlStr.Append("					<td colspan=8 class=xl14522200 width=424                                                                                 \n");
                                htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 318pt;border-bottom:none;'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                     \n");
                                htmlStr.Append("					<td colspan=2 class=xl14822200 width=49                                                                                  \n");
                                htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 37pt;border-bottom:none;'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                      \n");
                                htmlStr.Append("					<td colspan=2 class=xl15022200 width=79                                                                                  \n");
                                htmlStr.Append("						style='border-left: none; width: 60pt;border-bottom:none;'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                      \n");
                                htmlStr.Append("					<td colspan=2 class=xl15522200 width=87                                                                                  \n");
                                htmlStr.Append("						style='border-right: 1pt solid black; width: 65pt;border-bottom:none;'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                         \n");
                                htmlStr.Append("					<td colspan=2 class=xl15722200 width=123                                                                                 \n");
                                htmlStr.Append("						style='border-right: 2.0pt double black; border-left: none; width: 92pt;border-bottom:none;'>&nbsp;" + dt_d.Rows[v_index][4] + "&nbsp;</td>                    \n");
                                htmlStr.Append("				</tr>                                                                                                                        \n");

                            }

                        }
                    }
                    else
                    { // dong giua
                      // 
                        htmlStr.Append("				<tr height=24 style='mso-height-source: userset; height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>                                                           \n");
                        htmlStr.Append("					<td height=24 class=xl8922200                                                                                            \n");
                        htmlStr.Append("						style='height: " + (k == v_countNumberOfPages - 1 ? v_rowHeightLast : v_rowHeight).ToString() + "'>&nbsp;" + dt_d.Rows[v_index][7] + "</td>                                                   \n");
                        htmlStr.Append("					<td colspan=8 class=xl14522200 width=424                                                                                 \n");
                        htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 318pt'>&nbsp;" + dt_d.Rows[v_index]["ITEM_NAME"].ToString() + "</td>                     \n");
                        htmlStr.Append("					<td colspan=2 class=xl14822200 width=49                                                                                  \n");
                        htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 37pt'>&nbsp;" + dt_d.Rows[v_index][1] + "</td>                      \n");
                        htmlStr.Append("					<td colspan=2 class=xl15022200 width=79                                                                                  \n");
                        htmlStr.Append("						style='border-left: none; width: 60pt'>" + dt_d.Rows[v_index][2] + "&nbsp;</td>                                                      \n");
                        htmlStr.Append("					<td colspan=2 class=xl15022200 width=87                                                                                  \n");
                        htmlStr.Append("						style='border-right: 1pt solid black; width: 65pt'>" + dt_d.Rows[v_index][3] + "&nbsp;</td>                                         \n");
                        htmlStr.Append("					<td colspan=2 class=xl15922200 width=123                                                                                 \n");
                        htmlStr.Append("						style='border-right: 2.0pt double black; width: 92pt'>" + dt_d.Rows[v_index][4] + "&nbsp;</td>                                       \n");
                        htmlStr.Append("				</tr>                                                                                                                        \n");

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
                    v_spacePerPage = 10;
                }

                if (k == v_countNumberOfPages - 1 && page_index[k] < rowsPerPage) // Trang cuoi khong du dong
                {
                    v_rowHeightEmptyLast = Math.Round(v_totalHeightLastPage / (rowsPerPage - page_index[k]), 2).ToString() + "pt";
                    for (int i = 0; i < rowsPerPage - page_index[k]; i++)
                    {
                        if (i == (rowsPerPage - page_index[k] - 1))
                        {
                            htmlStr.Append("				<tr height=24 style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "'>                                                           \n");
                            htmlStr.Append("					<td height=24 class=xl8922200                                                                                            \n");
                            htmlStr.Append("						style='height: " + v_rowHeightEmptyLast + ";'>&nbsp;</td>                                                   \n");
                            htmlStr.Append("					<td colspan=8 class=xl14522200 width=424                                                                                 \n");
                            htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 318pt;'>&nbsp;</td>                     \n");
                            htmlStr.Append("					<td colspan=2 class=xl14822200 width=49                                                                                  \n");
                            htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 37pt;'>&nbsp;</td>                      \n");
                            htmlStr.Append("					<td colspan=2 class=xl15022200 width=79                                                                                  \n");
                            htmlStr.Append("						style='border-left: none; width: 60pt;'>&nbsp;</td>                                                      \n");
                            htmlStr.Append("					<td colspan=2 class=xl15522200 width=87                                                                                  \n");
                            htmlStr.Append("						style='border-right: 1pt solid black; width: 65pt;'>&nbsp;</td>                                         \n");
                            htmlStr.Append("					<td colspan=2 class=xl15722200 width=123                                                                                 \n");
                            htmlStr.Append("						style='border-right: 2.0pt double black; border-left: none; width: 92pt;'>&nbsp;&nbsp;</td>                    \n");
                            htmlStr.Append("				</tr>                 \n");


                        }
                        else
                        {
                            htmlStr.Append("				<tr height=24 style='mso-height-source: userset; height: " + v_rowHeightEmptyLast + "'>                                                           \n");
                            htmlStr.Append("					<td height=24 class=xl8922200                                                                                            \n");
                            htmlStr.Append("						style='height: " + v_rowHeightEmptyLast + "; border-top: none'>&nbsp;</td>                                                                \n");
                            htmlStr.Append("					<td colspan=8 class=xl14822200 width=424                                                                                 \n");
                            htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 318pt'>&nbsp;</td>                                  \n");
                            htmlStr.Append("					<td colspan=2 class=xl14822200 width=49                                                                                  \n");
                            htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 37pt'>&nbsp;</td>                                   \n");
                            htmlStr.Append("					<td colspan=2 class=xl14822200 width=79                                                                                  \n");
                            htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 60pt'>&nbsp;</td>                                   \n");
                            htmlStr.Append("					<td colspan=2 class=xl14822200 width=87                                                                                  \n");
                            htmlStr.Append("						style='border-right: 1pt solid black; border-left: none; width: 65pt'>&nbsp;</td>                                   \n");
                            htmlStr.Append("					<td colspan=2 class=xl15922200 width=123                                                                                 \n");
                            htmlStr.Append("						style='border-right: 2.0pt double black; width: 92pt'>&nbsp;</td>                                                    \n");
                            htmlStr.Append("				</tr>                                                                                                                        \n");
                            htmlStr.Append("	                                                                                                                                         \n");
                        }
                    } // for

                }//Trang cuoi 11 dong

                if (k < v_countNumberOfPages - 1)
                {
                    htmlStr.Append("				<tr height=24 style='mso-height-source: userset; height: " + (v_spacePerPage).ToString() + "pt'>                                                           \n");
                    htmlStr.Append("					<td height=24 class=xl8922200                                                                                            \n");
                    htmlStr.Append("						style='height: " + (v_spacePerPage).ToString() + "pt;border-bottom: 1pt solid black;border-right: none;'>&nbsp;</td>                                                   \n");
                    htmlStr.Append("					<td colspan=8 class=xl14522200 width=424                                                                                 \n");
                    htmlStr.Append("						style='border-right: none; border-left: none;border-bottom: 1pt solid black; width: 318pt;'>&nbsp;</td>                     \n");
                    htmlStr.Append("					<td colspan=2 class=xl14822200 width=49                                                                                  \n");
                    htmlStr.Append("						style='border-right:none; border-left: none;border-bottom: 1pt solid black; width: 37pt;'>&nbsp;</td>                      \n");
                    htmlStr.Append("					<td colspan=2 class=xl15022200 width=79                                                                                  \n");
                    htmlStr.Append("						style='border-left: none;border-bottom: 1pt solid black;border-right:none;  width: 60pt;'>&nbsp;</td>                                                      \n");
                    htmlStr.Append("					<td colspan=2 class=xl15522200 width=87                                                                                  \n");
                    htmlStr.Append("						style='border-right:none;border-bottom: 1pt solid black;border-left: none; width: 65pt;'>&nbsp;</td>                                         \n");
                    htmlStr.Append("					<td colspan=2 class=xl15722200 width=123                                                                                 \n");
                    htmlStr.Append("						style='border-right: 2.0pt double black;border-bottom: 1pt solid black; border-left: none; width: 92pt;'>&nbsp;&nbsp;</td>                    \n");
                    htmlStr.Append("				</tr>                                                                                                                        \n");
                    //     htmlStr.Append("	<table  border=0>                                                                                                                                                                                                 \n");
                    //   htmlStr.Append("		<tr height=0  style='height: 0pt'>                                                                                                                                                                \n");
                    //
                    //   htmlStr.Append("		</tr>      																																														\n");
                    //  htmlStr.Append("	</table>             																																										\n");

                }


            }// for k                                                                                                                             

            htmlStr.Append("				<tr class=xl11822200 height=26                                                                                               \n");
            htmlStr.Append("					style='mso-height-source: userset; height: 19.95pt'>                                                                     \n");
            htmlStr.Append("					<td colspan=2 height=26 class=xl12522200 style='height: 19.95pt'><span                                                   \n");
            htmlStr.Append("						style='mso-spacerun: yes'> </span>T&#7927; giá / T&#7893;ng USD:</td>                                                \n");
            htmlStr.Append("					<td class=xl11622200>&nbsp;" + lb_amount_trans + "</td>                                                                        \n");
            htmlStr.Append("					<td colspan=3 class=xl11622200>&nbsp;</td>                                                                               \n");
            htmlStr.Append("					<td class=xl11722200>&nbsp;" + amount_trans + "</td>                                                                          \n");
            htmlStr.Append("					<td class=xl11722200>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("					<td class=xl11722200>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("					<td colspan=6 class=xl11922200                                                                                           \n");
            htmlStr.Append("						style='border-right: 1pt solid black'>C&#7897;ng ti&#7873;n                                                         \n");
            htmlStr.Append("						hàng <font class='font622200'>(Total):</font><font                                                                   \n");
            htmlStr.Append("						class='font522200'><span style='mso-spacerun: yes'>  </span></font>                                                  \n");
            htmlStr.Append("					</td>                                                                                                                    \n");
            htmlStr.Append("					<td colspan=2 class=xl12722200                                                                                           \n");
            htmlStr.Append("						style='border-right: 2.0pt double black; border-left: none'>" + amount_net + "&nbsp;</td>                               \n");
            htmlStr.Append("				</tr>                                                                                                                        \n");
            htmlStr.Append("				<tr class=xl11822200 height=26                                                                                               \n");
            htmlStr.Append("					style='mso-height-source: userset; height: 19.95pt'>                                                                     \n");
            htmlStr.Append("					<td colspan=2 height=26 class=xl12522200 style='height: 19.95pt'><span                                                   \n");
            htmlStr.Append("						style='mso-spacerun: yes'> </span>Thu&#7871; su&#7845;t <font                                                        \n");
            htmlStr.Append("						class='font622200'>(Rate)</font><font class='font522200'>:<span                                                      \n");
            htmlStr.Append("							style='mso-spacerun: yes'>                                  </span></font></td>                                  \n");
            htmlStr.Append("					<td class=xl11622200 style='border-top: none'>&nbsp;</td>                                                                \n");
            htmlStr.Append("					<td class=xl11922200 style='border-top: none'>" + dt.Rows[0]["taxrate"] + "</td>                                                      \n");
            htmlStr.Append("					<td class=xl11722200 style='border-top: none'>&nbsp;</td>                                                                \n");
            htmlStr.Append("					<td class=xl11722200 style='border-top: none'>&nbsp;</td>                                                                \n");
            htmlStr.Append("					<td class=xl12022200 style='border-top: none'>&nbsp;</td>                                                                \n");
            htmlStr.Append("					<td class=xl12022200 style='border-top: none'>&nbsp;</td>                                                                \n");
            htmlStr.Append("					<td class=xl12022200 style='border-top: none'>&nbsp;</td>                                                                \n");
            htmlStr.Append("					<td colspan=6 class=xl11922200                                                                                           \n");
            htmlStr.Append("						style='border-right: 1pt solid black'>Thu&#7871; GTGT <font                                                         \n");
            htmlStr.Append("						class='font622200'>(VAT):</font><font class='font522200'></font></td>                                                \n");
            htmlStr.Append("					<td colspan=2 class=xl12722200                                                                                           \n");
            htmlStr.Append("						style='border-right: 2.0pt double black; border-left: none;'>" + amount_vat + "&nbsp;</td>                              \n");
            htmlStr.Append("				</tr>                                                                                                                        \n");
            htmlStr.Append("				<tr class=xl11822200 height=26                                                                                               \n");
            htmlStr.Append("					style='mso-height-source: userset; height: 19.95pt'>                                                                     \n");
            htmlStr.Append("					<td colspan=6 height=26 class=xl12522200 style='height: 19.95pt'>&nbsp;</td>                                             \n");
            htmlStr.Append("					<td class=xl11722200>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("					<td class=xl11722200>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("					<td colspan=7 class=xl11922200                                                                                           \n");
            htmlStr.Append("						style='border-right: 1pt solid black'>T&#7893;ng s&#7889;                                                           \n");
            htmlStr.Append("						ti&#7873;n thanh toán <font class='font622200'>(Invoice                                                              \n");
            htmlStr.Append("							total):</font></td>                                                                                              \n");
            htmlStr.Append("					<td colspan=2 class=xl12722200                                                                                           \n");
            htmlStr.Append("						style='border-right: 2.0pt double black; border-left: none;font-weight: 700'>" + amount_total + "&nbsp;</td>     \n");
            htmlStr.Append("				</tr>                                                                                                                        \n");
            htmlStr.Append("				<tr height=24 style='mso-height-source: userset; height: 21.0pt'>                                                           \n");
            htmlStr.Append("					<td height=24 class=xl9122200 colspan=6 style='height: 21.0pt'><span                                                    \n");
            htmlStr.Append("						style='mso-spacerun: yes'> </span>S&#7889; ti&#7873;n vi&#7871;t                                                     \n");
            htmlStr.Append("						b&#7857;ng ch&#7919; <font class='font622200'>(Total amount                                                          \n");
            htmlStr.Append("							in word)</font><font class='font522200'>:<span                                                                   \n");
            htmlStr.Append("							style='mso-spacerun: yes'> </span></font></td>                                                                   \n");
            htmlStr.Append("					<td colspan=11 class=xl13222200                                                                                          \n");
            htmlStr.Append("						style='border-right: 2.0pt double black;font-size: 12.5pt; '>&nbsp;" + read_prive + "</td>                                                 \n");
            htmlStr.Append("				</tr>                                                                                                                        \n");
            htmlStr.Append("				<tr height=24 style='mso-height-source: userset; height: 21.0pt'>                                                           \n");
            htmlStr.Append("					<td colspan=17 height=24 class=xl12922200                                                                                \n");
            htmlStr.Append("						style='border-right: 2.0pt double black; height: 21.0pt'>&nbsp;</td>                                                \n");
            htmlStr.Append("				</tr>                                                                                                                        \n");
            htmlStr.Append("				<tr height=6 style='mso-height-source: userset; height: 2.25pt'>                                                             \n");
            htmlStr.Append("					<td colspan=17 height=6 class=xl19022200                                                                                 \n");
            htmlStr.Append("						style='border-right: 2.0pt double black; height: 2.25pt'>&nbsp;</td>                                                 \n");
            htmlStr.Append("				</tr>                                                                                                                        \n");
            htmlStr.Append("				<tr height=26 style='mso-height-source: userset; height: 20.25pt'>                                                           \n");
            htmlStr.Append("					<td colspan=8 height=26 class=xl13522200 style='height: 20.25pt'>Khách                                                   \n");
            htmlStr.Append("						hàng <font class='font822200'>(Customer)</font>                                                                      \n");
            htmlStr.Append("					</td>                                                                                                                    \n");
            htmlStr.Append("					<td colspan=1 class=xl13222200></td>                                                                                     \n");
            htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td colspan=5 class=xl13222200 style='text-align:center'><span style='mso-spacerun: yes'>                                \n");
            htmlStr.Append("					</span>Giám &#273;&#7889;c <font class='font822200'>(Director)</font></td>                                               \n");
            htmlStr.Append("					<td class=xl9222200 style='border-top: none'>&nbsp;</td>                                                                 \n");
            htmlStr.Append("				</tr>                                                                                                                        \n");
            htmlStr.Append("				<tr height=24 style='height: 21.0pt'>                                                                                       \n");
            htmlStr.Append("					<td colspan=8 height=24 class=xl13622200 style='height: 21.0pt'>(Ký,                                                    \n");
            htmlStr.Append("						ghi rõ h&#7885; tên)</td>                                                                                            \n");
            htmlStr.Append("					<td colspan=1 class=xl13322200></td>                                                                                     \n");
            htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td colspan=5 class=xl13322200><span style='mso-spacerun: yes'>                                                          \n");
            htmlStr.Append("					</span>(Ký,&#273;óng d&#7845;u,ghi rõ h&#7885; tên)</td>                                                                 \n");
            htmlStr.Append("					<td class=xl9422200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("				</tr>                                                                                                                        \n");
            htmlStr.Append("				<tr height=24 style='height: 21.0pt'>                                                                                       \n");
            htmlStr.Append("					<td colspan=8 height=24 class=xl14422200 style='height: 21.0pt'>(Signature,                                             \n");
            htmlStr.Append("						full name)</td>                                                                                                      \n");
            htmlStr.Append("					<td colspan=1 class=xl13422200></td>                                                                                     \n");
            htmlStr.Append("					<td class=xl9522200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td class=xl9522200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td colspan=5 class=xl13722200><span style='mso-spacerun: yes'>                                                          \n");
            htmlStr.Append("					</span>(Signature,seal &amp; full name)</td>                                                                             \n");
            htmlStr.Append("					<td class=xl9622200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("				</tr>                                                                                                                        \n");
            htmlStr.Append("    		<tr height=33 style='mso-height-source: userset; height: 20.05pt'>                                                               \n");
            htmlStr.Append("    				<td height=33 class=xl9722200 style='height: 20.05pt'>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9422200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    			</tr>                                                               \n");
            htmlStr.Append("    			<tr height=24 style='height: 15.30pt'>                                                               \n");
            htmlStr.Append("    				<td height=24 class=xl9722200 colspan=2 style='height: 15.30pt'></td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9822200 colspan=3>Signature Valid</span></td>                                                               \n");
            if (dt.Rows[0]["sign_yn"].ToString() == "Y")
            {

                htmlStr.Append("					<td align=left class=xl9922200 valign=top style='border-top: 1pt solid windowtext;'><![if !vml]><span                                   \n");
                htmlStr.Append("	style='mso-ignore: vglayout; position: absolute; z-index: 2; margin-left: 18px; margin-top: 7px; width: 80px; height : 54px'><img		 \n");
                htmlStr.Append("							width=80 height=54                                                                                               \n");
                htmlStr.Append("							src='D:\\webproject\\e-invoice-ws\\02.Web\\EInvoice\\img\\check_signed.png'                                             \n");
                htmlStr.Append("							v:shapes='Picture_x0020_8'></span>                                                                               \n");
                htmlStr.Append("					<![endif]><span style='mso-ignore: vglayout2'>                                                                           \n");
                htmlStr.Append("							<table cellpadding=0 cellspacing=0>                                                                              \n");
                htmlStr.Append("								<tr>                                                                                                         \n");
                htmlStr.Append("									<td height=24 width=38                                                                   \n");
                htmlStr.Append("										style='height: 22.0pt; width: 24.65pt;border-top: none'>&nbsp;</td>                                 \n");
                htmlStr.Append("								</tr>                                                                                                        \n");
                htmlStr.Append("							</table>                                                                                                         \n");
                htmlStr.Append("					</span></td>                                                                                                             \n");
            }
            else
            {
                htmlStr.Append("							<td height=24 class=xl9922200 width=38                                                                           \n");
                htmlStr.Append("										style='height: 22.0pt; width: 24.65pt'>&nbsp;</td>                                                  \n");
            }
            htmlStr.Append("    				<td class=xl10022200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl10022200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl10122200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl10222200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    			</tr>                                                               \n");
            htmlStr.Append("    			<tr height=24 style='height: 15.30pt'>                                                               \n");
            htmlStr.Append("    				<td height=24 class=xl10322200 style='height: 15.30pt'> </span>                                                               \n");
            htmlStr.Append("    				</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    				<td colspan=7 class=xl13822200                                                               \n");
            htmlStr.Append("    					style='border-right: 1pt solid black'><font                                                               \n");
            htmlStr.Append("    					class='font1122200'>&nbsp;&#272;&#432;&#7907;c ký b&#7903;i:</font><font                                                               \n");
            htmlStr.Append("    					class='font1022200'> " + dt.Rows[0]["SignedBy"] + "<span                                                               \n");
            htmlStr.Append("    						style='mso-spacerun: yes'> </span></font></td>                                                               \n");
            htmlStr.Append("    				<td class=xl10422200>&nbsp;</td>                                                               \n");
            htmlStr.Append("    			</tr>                                                               \n");
            htmlStr.Append("    			 <tr height=24 style='height:15.30pt'>                                                               \n");
            htmlStr.Append("      <td height=24 class=xl10322200 style='height:15.30pt'>&nbsp;Mã CQT: " + dt.Rows[0]["cqt_mccqt_id"] + "</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl10622200 colspan=5>Ngày Ký: <font class='font1322200'>" + dt.Rows[0]["SignedDate"] + "</font></td>                                                               \n");
            htmlStr.Append("      <td class=xl10722200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl10822200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl10922200>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            htmlStr.Append("      <tr height=33 style='mso-height-source:userset;height:15.05pt'>                                                               \n");
            htmlStr.Append("      <td height=33 class=xl9722200 style='height:15.05pt'>&nbsp;Tra c&#7913;u t &#7841;i Website: <font class='font722200'><span                                                               \n");
            htmlStr.Append("      style='mso-spacerun:yes'> </span></font><font class='font1222200'>" + dt.Rows[0]["WEBSITE_EI"] + "</font></span></td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>Mã nh&#7853;n hóa &#273;&#417;n: " + dt.Rows[0]["matracuu"] + "</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9322200>&nbsp;</td>                                                               \n");
            htmlStr.Append("      <td class=xl9422200>&nbsp;</td>                                                               \n");
            htmlStr.Append("     </tr>                                                               \n");
            /*
                        htmlStr.Append("				<tr height=33 style='mso-height-source: userset; height: 35.05pt'>                                                           \n");
                        htmlStr.Append("					<td height=33 class=xl9722200 style='height: 35.05pt'>&nbsp;</td>                                                        \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9422200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("				</tr>                                                                                                                        \n");
                        htmlStr.Append("				<tr height=24 style='height: 21.0pt'>                                                                                       \n");
                        htmlStr.Append("					<td height=24 class=xl9722200 colspan=2 style='height: 21.0pt'></td>                                                    \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9822200 colspan=3>Signature Valid</span></td>                                                                \n");

                        if (dt.Rows[0]["sign_yn"].ToString() == "Y")
                        {

                            htmlStr.Append("					<td align=left valign=top style='border-top: 1pt solid windowtext;'><![if !vml]><span                                   \n");
                            htmlStr.Append("	style='mso-ignore: vglayout; position: absolute; z-index: 2; margin-left: 18px; margin-top: 7px; width: 80px; height : 54px'><img		 \n");
                            htmlStr.Append("							width=80 height=54                                                                                               \n");
                            htmlStr.Append("							src='E:/Tomcat/webapps/e-invoice/assets/img/check_signed.png'                                             \n");
                            htmlStr.Append("							v:shapes='Picture_x0020_8'></span>                                                                               \n");
                            htmlStr.Append("					<![endif]><span style='mso-ignore: vglayout2'>                                                                           \n");
                            htmlStr.Append("							<table cellpadding=0 cellspacing=0>                                                                              \n");
                            htmlStr.Append("								<tr>                                                                                                         \n");
                            htmlStr.Append("									<td height=24 class=xl9922200 width=38                                                                   \n");
                            htmlStr.Append("										style='height: 22.0pt; width: 24.65pt;border-top: none'>&nbsp;</td>                                 \n");
                            htmlStr.Append("								</tr>                                                                                                        \n");
                            htmlStr.Append("							</table>                                                                                                         \n");
                            htmlStr.Append("					</span></td>                                                                                                             \n");
                        }
                        else
                        {
                            htmlStr.Append("							<td height=24 class=xl9922200 width=38                                                                           \n");
                            htmlStr.Append("										style='height: 22.0pt; width: 24.65pt'>&nbsp;</td>                                                  \n");
                        }
                        htmlStr.Append("					<td class=xl10022200>&nbsp;</td>                                                                                         \n");
                        htmlStr.Append("					<td class=xl10022200>&nbsp;</td>                                                                                         \n");
                        htmlStr.Append("					<td class=xl10122200>&nbsp;</td>                                                                                         \n");
                        htmlStr.Append("					<td class=xl10222200>&nbsp;</td>                                                                                         \n");
                        htmlStr.Append("				</tr>                                                                                                                        \n");
                        htmlStr.Append("				<tr height=24 style='height: 21.0pt'>                                                                                       \n");
                        htmlStr.Append("					<td height=24 class=xl10322200 style='height: 21.0pt'>&nbsp;Mã                                                          \n");
                        htmlStr.Append("						nh&#7853;n hóa &#273;&#417;n: " + dt.Rows[0]["matracuu"] + "</span>                                                                \n");
                        htmlStr.Append("					</td>                                                                                                                    \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td colspan=7 class=xl13822200                                                                                           \n");
                        htmlStr.Append("						style='border-right: 1pt solid black'><font                                                                         \n");
                        htmlStr.Append("						class='font1122200'>&nbsp;&#272;&#432;&#7907;c ký b&#7903;i:</font><font                                             \n");
                        htmlStr.Append("						class='font1022200'> " + dt.Rows[0]["SignedBy"] + "<span                                                                          \n");
                        htmlStr.Append("							style='mso-spacerun: yes'> </span></font></td>                                                                   \n");
                        htmlStr.Append("					<td class=xl10422200>&nbsp;</td>                                                                                         \n");
                        htmlStr.Append("				</tr>                                                                                                                        \n");
                        htmlStr.Append("				<tr height=24 style='height: 21.0pt'>                                                                                       \n");
                        htmlStr.Append("					<td height=24 class=xl10522200 style='height: 21.0pt'>&nbsp;Tra                                                         \n");
                        htmlStr.Append("						c&#7913;u t &#7841;i Website: <font class='font722200'><span                                                         \n");
                        htmlStr.Append("							style='mso-spacerun: yes'> </span></font><font class='font1222200'>" + dt.Rows[0]["WEBSITE_EI"] + "</font></span > \n");
                        htmlStr.Append("					</td>                                                                                                                    \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
                        htmlStr.Append("					<td class=xl10622200 colspan=5>Ngày Ký: <font                                                                            \n");
                        htmlStr.Append("						class='font1322200'>" + dt.Rows[0]["SignedDate"] + "</font></td>                                                                       \n");
                        htmlStr.Append("					<td class=xl10722200>&nbsp;</td>                                                                                         \n");
                        htmlStr.Append("					<td class=xl10822200>&nbsp;</td>                                                                                         \n");
                        htmlStr.Append("					<td class=xl10922200>&nbsp;</td>                                                                                         \n");
                        htmlStr.Append("				</tr>                                                                                                                        \n");

                        */
            htmlStr.Append("				<!-- <tr height=13 style='mso-height-source: userset; height: 10.05pt'>                                                      \n");
            htmlStr.Append("					<td height=13 class=xl10522200 style='height: 10.05pt'>&nbsp;</td>                                                       \n");
            htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td class=xl9322200>&nbsp;</td>                                                                                          \n");
            htmlStr.Append("					<td class=xl12122200>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("					<td class=xl12122200>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("					<td class=xl12122200>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("					<td class=xl12122200>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("					<td class=xl12122200>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("					<td class=xl12122200>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("					<td class=xl12122200>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("					<td class=xl10922200>&nbsp;</td>                                                                                         \n");
            htmlStr.Append("				</tr>-->                                                                                                                     \n");
            htmlStr.Append("				<tr height=20 style='mso-height-source: userset; height: 12.75pt'>                                                           \n");
            htmlStr.Append("					<td colspan=17 height=20 class=xl14122200                                                                                \n");
            htmlStr.Append("						style='border-right: 2.0pt double black; height: 12.75pt'>" + dt.Rows[0]["CONTRACT_INFO_EI"] + "                                                     \n");
            htmlStr.Append("					</td>                                                                                                                    \n");
            htmlStr.Append("				</tr>                                                                                                                        \n");
            htmlStr.Append("				<![if supportMisalignedColumns]>                                                                                             \n");
            htmlStr.Append("				<tr height=0 style='display: none'>                                                                                          \n");
            htmlStr.Append("					<td width=59 style='width: 37.4pt'></td>                                                                                 \n");
            htmlStr.Append("					<td width=107 style='width: 68pt'></td>                                                                                  \n");
            htmlStr.Append("					<td width=23 style='width: 14.5pt'></td>                                                                                 \n");
            htmlStr.Append("					<td width=46 style='width: 28.9pt'></td>                                                                                 \n");
            htmlStr.Append("					<td width=40 style='width: 30.0pt'></td>                                                                                 \n");
            htmlStr.Append("					<td width=29 style='width: 18.7pt'></td>                                                                                 \n");
            htmlStr.Append("					<td width=34 style='width: 22.1pt'></td>                                                                                 \n");
            htmlStr.Append("					<td width=25 style='width: 16.15pt'></td>                                                                                \n");
            htmlStr.Append("					<td width=120 style='width: 76.5pt'></td>                                                                                \n");
            htmlStr.Append("					<td width=19 style='width: 14.0pt'></td>                                                                                 \n");
            htmlStr.Append("					<td width=30 style='width: 19.5pt'></td>                                                                                 \n");
            htmlStr.Append("					<td width=41 style='width: 26.35pt'></td>                                                                                \n");
            htmlStr.Append("					<td width=38 style='width: 24.65pt'></td>                                                                                \n");
            htmlStr.Append("					<td width=49 style='width: 31.45pt'></td>                                                                                \n");
            htmlStr.Append("					<td width=38 style='width: 23.8pt'></td>                                                                                 \n");
            htmlStr.Append("					<td width=106 style='width: 67.15pt'></td>                                                                               \n");
            htmlStr.Append("					<td width=17 style='width: 11.05pt'></td>                                                                                \n");
            htmlStr.Append("				</tr>                                                                                                                        \n");
            htmlStr.Append("				<![endif]>                                                                                                                   \n");
            htmlStr.Append("			</table>                                                                                                                         \n");
            htmlStr.Append("	                                                                                                                                         \n");
            htmlStr.Append("		</div>                                                                                                                               \n");
            htmlStr.Append("	</body>                                                                                                                                  \n");
            htmlStr.Append("	</html>                                                                                                                                  \n");


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
            int result = 1;
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