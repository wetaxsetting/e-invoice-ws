
using iTextSharp.text.pdf;
using System;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.IO;
using System.Text;
using System.Xml;

namespace EInvoice.Company
{
    public class Demo
    {
        public static string View(string tei_einvoice_m_pk, string tei_company_pk, string dbName)
        {
            //dbName = "NOBLANDBD";
            string dbUser = "genuwin", dbPwd = "genuwin2";//NOBLANDBD  EINVOICE_252
            string _conString = "Data Source={0};User Id={1};Password={2};Unicode=true";
            _conString = String.Format(_conString, dbName, dbUser, dbPwd);

            string Procedure = "stacfdstac71_r_02_1";
            OracleConnection connection;
            connection = new OracleConnection(_conString);
            connection.Open();
            OracleCommand command = new OracleCommand(Procedure, connection);
            command.CommandType = CommandType.StoredProcedure;

            command.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = tei_einvoice_m_pk;
            command.Parameters.Add("p_user_id", OracleType.VarChar, 1000).Value = "genuwin";
            command.Parameters.Add("p_rtn_value", OracleType.Cursor).Direction = ParameterDirection.Output;
            DataSet ds = new DataSet();

            OracleDataAdapter da = new OracleDataAdapter(command);
            da.Fill(ds);
            DataTable dt = ds.Tables[0];

            command.Parameters.Clear();

            Procedure = "stacfdstac71_r_03";

            command = new OracleCommand(Procedure, connection);
            command.CommandType = CommandType.StoredProcedure;


            command.Parameters.Add("p_tei_einvoice_m_pk", OracleType.VarChar, 1000).Value = tei_einvoice_m_pk;
            command.Parameters.Add("p_user_id", OracleType.VarChar, 1000).Value = "genuwin";
            command.Parameters.Add("p_rtn_value", OracleType.Cursor).Direction = ParameterDirection.Output;
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
            if (dt.Rows[0]["CurrencyCodeUSD"].ToString() == "VND")
            {
                read_prive = NumberToTextVN(Decimal.Parse(dt.Rows[0]["TotalAmount"].ToString()));
            }
            else
            {
                read_prive = Num2VNText(dt.Rows[0]["TotalAmountInWord"].ToString().ToString(), "USD");
            }
            //read_prive = NumberToTextVN(Total_Amount_d);
            read_prive = read_prive.Replace(",", "");

            read_prive = read_prive.ToString().Substring(0, 2) + read_prive.ToString().Substring(2, read_prive.Length - 2).ToLower().Replace("mỹ", "Mỹ");

            read_prive = read_prive.Substring(0, 2) + read_prive.Substring(2, read_prive.Length - 2).ToLower() + ".";

            int end = 0;
            int count = count_page_v + r;
            double height = 130;
            StringBuilder htmlStr = new StringBuilder("");
            string heigh = "", heigh_d = "";

            htmlStr.Append("<!DOCTYPE html PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>																						\n");
            htmlStr.Append("<html>                                                                                                                                                                                      \n");
            htmlStr.Append("<head>                                                                                                                                                                                      \n");
            htmlStr.Append("<meta http-equiv='Content-Type' content='text/html; charset=UTF-8'>                                                                                                                         \n");
            htmlStr.Append("                                                                                                                                                                                            \n");
            htmlStr.Append("<script type='text/javascript' src='${pageContext.request.contextPath}/system/syscommand.js'></script>                                                                                      \n");
            htmlStr.Append("<title>Report E-Invoice</title>                                                                                                                                                             \n");
            //htmlStr.Append(" <!-- Normalize or reset CSS with your favorite library -->                                                                                                                                 \n");
            //htmlStr.Append(" <!-- <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/normalize/3.0.3/normalize.css'> -->                                                                                      \n");
            //htmlStr.Append("                                                                                                                                                                                            \n");
            //htmlStr.Append("  <!-- Load paper.css for happy printing -->                                                                                                                                                \n");
            //htmlStr.Append("  <!--<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/paper-css/0.2.3/paper.css'>  -->                                                                                         \n");
            //htmlStr.Append("                                                                                                                                                                                            \n");
            //htmlStr.Append("  <!-- Set page size here: A5, A4 or A3 -->                                                                                                                                                 \n");
            //htmlStr.Append("  <!-- Set also 'landscape' if you need -->                                                                                                                                                 \n");
            //htmlStr.Append("  <style>@page { size: A4 }</style>                                                                                                                                                         \n");
            //htmlStr.Append("  <!--<link href='https://fonts.googleapis.com/css?family=Tangerine:700' rel='stylesheet' type='text/css'>  -->                                                                                    \n");
            htmlStr.Append("  <style>                                                                                                                                                                                   \n");
            htmlStr.Append("    /*body   { font-family: serif }                                                                                                                                                         \n");
            htmlStr.Append("    h1     { font-family: 'Tangerine', cursive; font-size: 40pt; line-height: 18mm}                                                                                                         \n");
            htmlStr.Append("    h2, h3 { font-family: 'Tangerine', cursive; font-size: 24pt; line-height: 7mm }                                                                                                         \n");
            htmlStr.Append("    h4     { font-size: 13pt; line-height: 1mm }                                                                                                                                            \n");
            htmlStr.Append("    h2 + p { font-size: 18pt; line-height: 7mm }                                                                                                                                            \n");
            htmlStr.Append("    h3 + p { font-size: 14pt; line-height: 7mm }                                                                                                                                            \n");
            htmlStr.Append("    li     { font-size: 11pt; line-height: 5mm }                                                                                                                                            \n");
            htmlStr.Append("                                                                                                                                                                                            \n");
            htmlStr.Append("    h1      { margin: 0 }                                                                                                                                                                   \n");
            htmlStr.Append("    h1 + ul { margin: 2mm 0 5mm }                                                                                                                                                           \n");
            htmlStr.Append("    h2, h3  { margin: 0 3mm 3mm 0; float: left }                                                                                                                                            \n");
            htmlStr.Append("    h2 + p,                                                                                                                                                                                 \n");
            htmlStr.Append("    h3 + p  { margin: 0 0 3mm 50mm }                                                                                                                                                        \n");
            htmlStr.Append("    //h4      { margin: 1mm 0 0 2mm; border-bottom: 1px solid black }                                                                                                                       \n");
            htmlStr.Append("    h4 + ul { margin: 5mm 0 0 50mm }                                                                                                                                                        \n");
            htmlStr.Append("    article { border: 4px double black; padding: 5mm 10mm; border-radius: 3mm }*/                                                                                                           \n");
            htmlStr.Append("    body {                                                                                                                                                                                  \n");
            htmlStr.Append("       		 color: blue;                                                                                                                                                                   \n");
            htmlStr.Append("       		 font-size:100%;                                                                                                                                                                \n");
            htmlStr.Append("       		 background-image: url('assets/Solution.jpg');                                                                                                                                  \n");
            htmlStr.Append("		 }                                                                                                                                                                                  \n");
            htmlStr.Append("	h1 {                                                                                                                                                                                    \n");
            htmlStr.Append("	        color: #00FF00;                                                                                                                                                                 \n");
            htmlStr.Append("	}                                                                                                                                                                                       \n");
            htmlStr.Append("	p {                                                                                                                                                                                     \n");
            htmlStr.Append("	        color: rgb(0,0,255)                                                                                                                                                             \n");
            htmlStr.Append("	}                                                                                                                                                                                       \n");
            htmlStr.Append("	                                                                                                                                                                                        \n");
            htmlStr.Append("   headline1 {                                                                                                                                                                              \n");
            htmlStr.Append("      background-image: url(assets/Solution.jpg);                                                                                                                                           \n");
            htmlStr.Append("      background-repeat: no-repeat;                                                                                                                                                         \n");
            htmlStr.Append("      background-position: left top;                                                                                                                                                        \n");
            htmlStr.Append("      padding-top:68px;                                                                                                                                                                     \n");
            htmlStr.Append("      margin-bottom:50px;                                                                                                                                                                   \n");
            htmlStr.Append("   }                                                                                                                                                                                        \n");
            htmlStr.Append("   headline2 {                                                                                                                                                                              \n");
            htmlStr.Append("      background-image: url(images/newsletter_headline2.gif);                                                                                                                               \n");
            htmlStr.Append("      background-repeat: no-repeat;                                                                                                                                                         \n");
            htmlStr.Append("      background-position: left top;                                                                                                                                                        \n");
            htmlStr.Append("      padding-top:68px;                                                                                                                                                                     \n");
            htmlStr.Append("   }                                                                                                                                                                                        \n");
            htmlStr.Append("<!--table                                                                                                                                                                                \n");
            htmlStr.Append("	{mso-displayed-decimal-separator:'\\.';                                                                                                                                                  \n");
            htmlStr.Append("	mso-displayed-thousand-separator:'\\,';}                                                                                                                                                 \n");
            htmlStr.Append("@page                                                                                                                                                                                       \n");
            htmlStr.Append("	{margin:.5in .3in .25in .5in;                                                                                                                                                           \n");
            htmlStr.Append("	mso-header-margin:.25in;                                                                                                                                                                \n");
            htmlStr.Append("	mso-footer-margin:.25in;}                                                                                                                                                               \n");
            htmlStr.Append("-->                                                                                                                                                                                         \n");
            htmlStr.Append("</style>                                                                                                                                                                                    \n");
            htmlStr.Append("                                                                                                                                                                                            \n");
            htmlStr.Append("</head>                                                                                                                                                                                     \n");
            htmlStr.Append("<body link='#0066CC' vlink=purple class=xl70>                                                                                                                                           \n");
            htmlStr.Append("                                                                                                                                                                                            \n");
            htmlStr.Append("<table x:str border=0 cellpadding=0 cellspacing=0 width=713 style='border-collapse:                                                                                                         \n");
            htmlStr.Append(" collapse;table-layout:fixed;width:536pt'>                                                                                                                                                  \n");
            htmlStr.Append(" <col class=xl70 width=7 style='mso-width-source:userset;mso-width-alt:256;                                                                                                                 \n");
            htmlStr.Append(" width:6.25pt'>                                                                                                                                                                                \n");
            htmlStr.Append(" <col class=xl70 width=32 style='mso-width-source:userset;mso-width-alt:1170;                                                                                                               \n");
            htmlStr.Append(" width:30pt'>                                                                                                                                                                               \n");
            htmlStr.Append(" <col class=xl70 width=159 style='mso-width-source:userset;mso-width-alt:5814;                                                                                                              \n");
            htmlStr.Append(" width:148.75pt'>                                                                                                                                                                              \n");
            htmlStr.Append(" <col class=xl70 width=80 style='mso-width-source:userset;mso-width-alt:2925;                                                                                                               \n");
            htmlStr.Append(" width:75pt'>                                                                                                                                                                               \n");
            htmlStr.Append(" <col class=xl70 width=61 style='mso-width-source:userset;mso-width-alt:2230;                                                                                                               \n");
            htmlStr.Append(" width:57.5pt'>                                                                                                                                                                               \n");
            htmlStr.Append(" <col class=xl70 width=54 style='mso-width-source:userset;mso-width-alt:1974;                                                                                                               \n");
            htmlStr.Append(" width:51.25pt'>                                                                                                                                                                               \n");
            htmlStr.Append(" <col class=xl70 width=61 style='mso-width-source:userset;mso-width-alt:2230;                                                                                                               \n");
            htmlStr.Append(" width:57.5pt'>                                                                                                                                                                               \n");
            htmlStr.Append(" <col class=xl70 width=33 style='mso-width-source:userset;mso-width-alt:1206;                                                                                                               \n");
            htmlStr.Append(" width:31.25pt'>                                                                                                                                                                               \n");
            htmlStr.Append(" <col class=xl70 width=82 style='mso-width-source:userset;mso-width-alt:2998;                                                                                                               \n");
            htmlStr.Append(" width:77.5pt'>                                                                                                                                                                               \n");
            htmlStr.Append(" <col class=xl70 width=12 style='mso-width-source:userset;mso-width-alt:438;                                                                                                                \n");
            htmlStr.Append(" width:11.25pt'>                                                                                                                                                                                \n");
            htmlStr.Append(" <col class=xl70 width=124 style='mso-width-source:userset;mso-width-alt:4534;                                                                                                              \n");
            htmlStr.Append(" width:116.25pt'>                                                                                                                                                                               \n");
            htmlStr.Append(" <col class=xl70 width=8 style='mso-width-source:userset;mso-width-alt:292;                                                                                                                 \n");
            htmlStr.Append(" width:7.5pt'>                                                                                                                                                                                \n");
            htmlStr.Append(" <tr class= height=20 style='height:347pt'>                                                                                                                                           \n");
            htmlStr.Append(" <td>                                                                                                                                                                                       \n");
            htmlStr.Append(" <![if !vml]><span style='mso-ignore:vglayout;position:                                                                                                               \n");
            htmlStr.Append("  absolute;z-index:4;margin-left:0px;margin-top:0px;width:891px;height:347px'><img                                                                                                          \n");
            htmlStr.Append("  width=891 height=347 src='C:/Users/genuwin/Desktop/1.png' v:shapes='Rectangle_x0020_1'></span><![endif]>                                                        \n");
            htmlStr.Append("  </td>                                                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                                      \n");
            htmlStr.Append(" <tr class=xl104 height=20 style='height:347pt'>                                                                                                                                           \n");
            htmlStr.Append(" <td>                                                                                                                                                                                       \n");
            htmlStr.Append(" <![if !vml]><span style='mso-ignore:vglayout;position:                                                                                                               \n");
            htmlStr.Append("  absolute;z-index:4;margin-left:0px;margin-top:0px;width:891px;height:347px'><img                                                                                                          \n");
            htmlStr.Append("  width=891 height=347 src='C:/Users/genuwin/Desktop/2.png' v:shapes='Rectangle_x0020_1'></span><![endif]>                                                        \n");
            htmlStr.Append("  </td>                                                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                                      \n");
            htmlStr.Append(" <tr class=xl104 height=20 style='height:347pt'>                                                                                                                                           \n");
            htmlStr.Append(" <td>                                                                                                                                                                                       \n");
            htmlStr.Append(" <![if !vml]><span style='mso-ignore:vglayout;position:                                                                                                               \n");
            htmlStr.Append("  absolute;z-index:4;margin-left:0px;margin-top:0px;width:891px;height:347px'><img                                                                                                          \n");
            htmlStr.Append("  width=891 height=347 src='C:/Users/genuwin/Desktop/3.png' v:shapes='Rectangle_x0020_1'></span><![endif]>                                                        \n");
            htmlStr.Append("  </td>                                                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                                      \n");
            htmlStr.Append(" <tr class=xl104 height=20 style='height:347pt'>                                                                                                                                           \n");
            htmlStr.Append(" <td>                                                                                                                                                                                       \n");
            htmlStr.Append(" <![if !vml]><span style='mso-ignore:vglayout;position:                                                                                                               \n");
            htmlStr.Append("  absolute;z-index:4;margin-left:0px;margin-top:0px;width:891px;height:347px'><img                                                                                                          \n");
            htmlStr.Append("  width=891 height=347 src='C:/Users/genuwin/Desktop/4.png' v:shapes='Rectangle_x0020_1'></span><![endif]>                                                        \n");
            htmlStr.Append("  </td>                                                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                                      \n");
            htmlStr.Append(" <tr class=xl104 height=20 style='height:347pt'>                                                                                                                                           \n");
            htmlStr.Append(" <td>                                                                                                                                                                                       \n");
            htmlStr.Append(" <![if !vml]><span style='mso-ignore:vglayout;position:                                                                                                               \n");
            htmlStr.Append("  absolute;z-index:4;margin-left:0px;margin-top:0px;width:891px;height:347px'><img                                                                                                          \n");
            htmlStr.Append("  width=891 height=347 src='C:/Users/genuwin/Desktop/5.png' v:shapes='Rectangle_x0020_1'></span><![endif]>                                                        \n");
            htmlStr.Append("  </td>                                                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                                      \n");
            htmlStr.Append(" <tr class=xl104 height=20 style='height:347pt'>                                                                                                                                           \n");
            htmlStr.Append(" <td>                                                                                                                                                                                       \n");
            htmlStr.Append(" <![if !vml]><span style='mso-ignore:vglayout;position:                                                                                                               \n");
            htmlStr.Append("  absolute;z-index:4;margin-left:0px;margin-top:0px;width:891px;height:347px'><img                                                                                                          \n");
            htmlStr.Append("  width=891 height=347 src='C:/Users/genuwin/Desktop/6.png' v:shapes='Rectangle_x0020_1'></span><![endif]>                                                        \n");
            htmlStr.Append("  </td>                                                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                                      \n");
            htmlStr.Append(" <tr class=xl104 height=20 style='height:347pt'>                                                                                                                                           \n");
            htmlStr.Append(" <td>                                                                                                                                                                                       \n");
            htmlStr.Append(" <![if !vml]><span style='mso-ignore:vglayout;position:                                                                                                               \n");
            htmlStr.Append("  absolute;z-index:4;margin-left:0px;margin-top:0px;width:891px;height:347px'><img                                                                                                          \n");
            htmlStr.Append("  width=891 height=347 src='C:/Users/genuwin/Desktop/7.png' v:shapes='Rectangle_x0020_1'></span><![endif]>                                                        \n");
            htmlStr.Append("  </td>                                                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                                      \n");
            htmlStr.Append(" <tr class=xl104 height=20 style='height:347pt'>                                                                                                                                           \n");
            htmlStr.Append(" <td>                                                                                                                                                                                       \n");
            htmlStr.Append(" <![if !vml]><span style='mso-ignore:vglayout;position:                                                                                                               \n");
            htmlStr.Append("  absolute;z-index:4;margin-left:0px;margin-top:0px;width:891px;height:347px'><img                                                                                                          \n");
            htmlStr.Append("  width=891 height=347 src='C:/Users/genuwin/Desktop/8.png' v:shapes='Rectangle_x0020_1'></span><![endif]>                                                        \n");
            htmlStr.Append("  </td>                                                                                                                                                                                     \n");
            htmlStr.Append(" </tr>                                                                                                                                                                                      \n");
            htmlStr.Append(" <![if supportMisalignedColumns]>                                                                                                                                                           \n");
            htmlStr.Append(" <tr height=0 style='display:none'>                                                                                                                                                         \n");
            htmlStr.Append("  <td width=7 style='width:6.25pt'></td>                                                                                                                                                       \n");
            htmlStr.Append("  <td width=32 style='width:30pt'></td>                                                                                                                                                     \n");
            htmlStr.Append("  <td width=159 style='width:148.75pt'></td>                                                                                                                                                   \n");
            htmlStr.Append("  <td width=80 style='width:75pt'></td>                                                                                                                                                     \n");
            htmlStr.Append("  <td width=61 style='width:57.5pt'></td>                                                                                                                                                     \n");
            htmlStr.Append("  <td width=54 style='width:51.25pt'></td>                                                                                                                                                     \n");
            htmlStr.Append("  <td width=61 style='width:57.5pt'></td>                                                                                                                                                     \n");
            htmlStr.Append("  <td width=33 style='width:31.25pt'></td>                                                                                                                                                     \n");
            htmlStr.Append("  <td width=82 style='width:77.5pt'></td>                                                                                                                                                     \n");
            htmlStr.Append("  <td width=12 style='width:11.25pt'></td>                                                                                                                                                      \n");
            htmlStr.Append("  <td width=124 style='width:116.25pt'></td>                                                                                                                                                    \n");
            htmlStr.Append("  <td width=8 style='width:7.5pt'></td>                                                                                                                                                       \n");
            htmlStr.Append(" </tr>                                                                                                                                                                                      \n");
            htmlStr.Append(" <![endif]>                                                                                                                                                                                 \n");
            htmlStr.Append("</table>                                                                                                                                                                                    \n");
            htmlStr.Append("</body>                                                                                                                                                                                     \n");
            htmlStr.Append("</html>                                                                                                                                                                                     \n");
            //string filePath = "C:/Users/genuwin/Desktop/" + tei_einvoice_m_pk + ".xml";
            //string filePath = "C:\\Users\\genuwin\\Desktop\\" + tei_einvoice_m_pk + ".xml";
            //string filePath = "E:\\webproject\\E_INVOICE_WS\\02.Web\\AttachFileXml\\" + tei_einvoice_m_pk + ".xml";


            // insert xml in database 


            return htmlStr.ToString() + "|" + "DEMO" + "_" + "MM" + "_" + "NO";
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