using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.IO;
using System.Drawing;
using NativeExcel;

public partial class stacrpstac710017 : System.Web.UI.Page
{
    
	protected void Page_Load(object sender, EventArgs e)
    {
		//ESysLib.SetUser(Session["APP_DBUSER"].ToString());
        //ESysLib.SetUser("acnt");
		
		string l_company_pk		= Request["p_company_pk"];			
		string l_Fiscal			= Request["p_Fiscal"];
		string l_from_date		= Request["p_from_date"];
		string l_to_date		= Request["p_to_date"];
		string l_Customer_CD	= Request["p_Customer_CD"];
		string l_BizPlace		= Request["p_BizPlace"];
		string l_FormNo			= Request["p_FormNo"];
		string l_EIStatus		= Request["p_EIStatus"];
		string l_TaxRate		= Request["p_TaxRate"];
		string l_TaxEInvoiceNo	= Request["p_TaxEInvoiceNo"];
		string l_InvoiceType	= Request["p_InvoiceType"];
		string l_SerialNo		= Request["p_SerialNo"];
		string l_EInvoiceNo		= Request["p_EInvoiceNo"];
		string l_Decision		= Request["p_Decision"];
		string l_TaxKey			= Request["p_TaxKey"];

        string TemplateFile = "stacrpstac710017.xlsx";
        string TempFile = "../../../system/temp/stacrpstac710017.xlsx";
        TemplateFile = Server.MapPath(TemplateFile);
        TempFile = Server.MapPath(TempFile);

        //Create a new workbook
        IWorkbook exBook = NativeExcel.Factory.OpenWorkbook(TemplateFile);

        //Add worksheet
        IWorksheet exSheet = exBook.Worksheets[1];
		//IWorksheet exSheet2 = exBook.Worksheets[2];

        //bind data to excel file
        string para = "";
        DataTable dt,dt_info;
        //para = "'" + l_tco_company_pk + "','" + l_tco_buspartner_pk + "','" + l_tr_date_fr + "','" + l_tr_date_to + "','" + l_tr_status + "','" + l_tr_type + "','" + l_tac_hgtrh_pk +"','" + l_voucherno + "','" + l_invoice_no + "','" + l_Item_pk + "','" + l_PLUnit + "','" + l_Nation + "','" + l_Form_No + "','" +l_SerialNo + "','" +l_InvoiceType +"','"+ l_EIStatus +"'";
	    //dt = ESysLib.TableReadOpenCursor("rpt_60110430", para);
			
		string SQL
          = 	"   SELECT  ROWNUM stt, 																					"+	
				"			z.form_no,                                                                                      "+
				"			z.serial_no,                                                                                    "+
				"			z.sign_dt,                                                                                      "+
				"			z.invoice_no,                                                                                   "+
				"			z.ei_status,                                                                                    "+
				"			z.customer_cd,                                                                                  "+
				"			z.customer_nm,                                                                                  "+
				"			z.tax_code,                                                                                     "+
				"			z.tot_net_tr_amt,                                                                               "+
				"			z.tax_rate,                                                                                     "+
				"			z.vat_amount,                                                                                   "+
				"			z.remark,                                                                                       "+
				"			z.remark2,                                                                                      "+
				"			z.pk                                                                                            "+
				"        FROM (  SELECT                                                                                     "+
				"                                                                                                           "+
				"                       a.form_no,                                                                          "+
				"                       a.serial_no,                                                                        "+
				"                       TO_CHAR (A.sign_dt, 'dd/mm/yyyy') sign_dt,                                          "+
				"                       a.invoice_no,                                                                       "+
				"                       (SELECT b.code_nm name                                                              "+
				"                          FROM tac_commcode_detail b, tac_commcode_master a                                "+
				"                         WHERE     a.del_if = 0                                                            "+
				"                               AND b.del_if = 0                                                            "+
				"                               AND B.USE_YN = 'Y'                                                          "+
				"                               AND b.tac_commcode_master_pk = a.pk                                         "+
				"                               AND a.id = 'ACEI0010'                                                       "+
				"                               AND B.CODE = A.ei_status)                                                   "+
				"                          ei_status,                                                                       "+
				"                       a.customer_cd,                                                                      "+
				"                       a.customer_nm,                                                                      "+
				"                       a.tax_code,                                                                         "+
				"                       a.TOT_NET_BK_AMT tot_net_tr_amt,                                                    "+
				"                       CASE                                                                                "+
				"                          WHEN    a.vat_rate = 'none'                                                      "+
				"                               OR a.vat_rate = 'NO'                                                        "+
				"                               OR a.vat_rate = '01'                                                        "+
				"                          THEN                                                                             "+
				"                             '-'                                                                           "+
				"                          WHEN a.vat_rate = '00'                                                           "+
				"                          THEN                                                                             "+
				"                             '0%'                                                                          "+
				"                          ELSE                                                                             "+
				"                             a.vat_rate || '%'                                                             "+
				"                       END                                                                                 "+
				"                          tax_rate,                                                                        "+
				"                       CASE                                                                                "+
				"                          WHEN     a.tot_vat_bk_amt = 0                                                    "+
				"                               AND (   a.vat_rate = 'none'                                                 "+
				"                                    OR a.vat_rate = 'NO'                                                   "+
				"                                    OR a.vat_rate = '01')                                                  "+
				"                          THEN                                                                             "+
				"                             0                                                                             "+
				"                          ELSE                                                                             "+
				"                             --(a.tot_net_bk_amt * TO_NUMBER (a.vat_rate)) / 100 lay tu table len          "+
				"                             a.tot_vat_tr_amt                                                              "+
				"                       END                                                                                 "+
				"                          vat_amount,                                                                      "+
				"                       a.remark,                                                                           "+
				"                       a.remark2,                                                                          "+
				"                        a.pk                                                                               "+
				"                  FROM tei_einvoice_m a                   --, tei_einvoice_d b                             "+
				"                 WHERE     a.del_if = 0                                                                    "+
				"                       --                       AND b.del_if = 0                                           "+
				"                       --                       AND a.pk = B.TEI_EINVOICE_M_PK                             "+
				"                       AND (   UPPER (a.customer_cd) LIKE                                                  "+
				"                                  '%' || UPPER ("+	l_Customer_CD	+") || '%'                              "+
				"                            OR UPPER (a.customer_nm) LIKE                                                  "+
				"                                  '%' || UPPER ("+	l_Customer_CD	+") || '%'                              "+
				"                            OR "+	l_Customer_CD	+" IS NULL)                                             "+
				"                       AND A.FORM_NO = "+	l_FormNo	+"                                                  "+
				"                       AND A.SERIAL_NO = "+	l_SerialNo	+"                                              "+
				"                       AND A.EI_STATUS in ('1','5')                                                        "+
				"                       AND (   UPPER (A.EI_STATUS) LIKE                                                    "+
				"                          '%' || UPPER (TRIM ("+	l_EIStatus	+")) || '%'                                 "+
				"                    OR "+	l_EIStatus	+" IS NULL)                                                         "+
				"                       AND A.TR_DATE BETWEEN "+	l_from_date	+" AND "+	l_to_date	+"                  "+
				"                       and A.TEI_COMPANY_PK = "+	l_company_pk	+"                                      "+
				"						--          				and A.INVOICE_TYPE = "+	l_InvoiceType	+"              "+
				"                        AND (   UPPER (A.INVOICE_TYPE) LIKE                                                "+
				"                          '%' || UPPER (TRIM ("+	l_InvoiceType	+" )) || '%'                            "+
				"                    OR "+	l_InvoiceType	+"  IS NULL)                                                    "+
				"                        AND (   UPPER (A.VAT_RATE) LIKE                                                    "+
				"                          '%' || UPPER (TRIM ("+	l_TaxRate	+" )) || '%'                                "+
				"                    OR "+	l_TaxRate	+" IS NULL)                                                         "+
				"                       --and A.VAT_RATE = "+	l_TaxRate	+"                                              "+
				"              ORDER BY  a.invoice_no ASC) z                                                                "+
				"          ORDER BY  stt;                                                                                   ";

		 
		// Response.Write(SQL);
        // Response.End();
        dt = ESysLib.TableReadOpen(SQL);
		//para = "'" + l_tco_company_pk +"'";
	    //dt_info = ESysLib.TableReadOpenCursor("ac_rpt_60110420_01_mst", para);
        //exSheet.Range["A1"].Value = dt_info.Rows[0]["PARTNER_LNAME"]; 
		//string strPrintTime = "Print date : " + DateTime.Now.ToString("dd/MM/yyyy hh:mm");
		//exSheet.Range["A2"].Value = strPrintTime;  
		//string v_status = "", v_invoice_type = "";	
		//exSheet.Range["C4"].Value = l_tr_date_fr.Substring(4,2) +"/"+l_tr_date_fr.Substring(6,2) +"/"+ l_tr_date_fr.Substring(0,4) +" ~ "+ l_tr_date_to.Substring(4,2) +"/"+l_tr_date_to.Substring(6,2) +"/"+ l_tr_date_to.Substring(0,4);
        //if (l_EIStatus == "1")
		//{
		//	v_status = "Issues";
		//}else if (l_EIStatus == "5")
		//{
		//	v_status = "Cancel";
		//}else 
		//{
		//	v_status = "Select All";
		//}
		
		
		
		//if (l_tr_type == "DO")
		//{
		//	v_invoice_type = "Domestic";
		//}else 
		//{
		//	v_invoice_type = "Foreign";
		//}
		
		//exSheet.Range["K4"].Value = v_status;  
		//exSheet.Range["O4"].Value = v_invoice_type;
		
		if (dt.Rows.Count  > 0)
        {
            Int32 start_row = 9;
            Int32 end_row = 8;
			Int32 v_tr_amt = 0, v_bk_amt = 0, v_vat_tr_amt = 0, v_vat_bk_amt = 0, v_tot_tr_amt = 0, v_tot_bk_amt = 0;
            for (int l_addrow = 0; l_addrow < dt.Rows.Count-1; l_addrow++)
            {
                exSheet.Range["A10"].Rows.EntireRow.Insert();//insert row new of sheet
            }
         //   exSheet.Cells[7, 3].Value = dt_info.Rows[0]["from_date1"].ToString();
            for (int i = 0; i < dt.Rows.Count ; i++)
            {
				exSheet.Cells[start_row + i, 1].Value = dt.Rows[i][0].ToString();
                exSheet.Cells[start_row + i, 2].Value = dt.Rows[i]["voucherno"].ToString();
                exSheet.Cells[start_row + i, 3].Value = "";//dt.Rows[i]["tr_date"].ToString();
				exSheet.Cells[start_row + i, 4].Value = dt.Rows[i]["form_no"].ToString();
				exSheet.Cells[start_row + i, 5].Value = dt.Rows[i]["serial_no"].ToString();
				exSheet.Cells[start_row + i, 6].Value = dt.Rows[i]["invoice_date"].ToString();
                exSheet.Cells[start_row + i, 7].Value = dt.Rows[i]["invoice_no"].ToString();
                exSheet.Cells[start_row + i, 8].Value = dt.Rows[i]["e_status"];
                exSheet.Cells[start_row + i, 9].Value = dt.Rows[i]["invoice_type"];
                exSheet.Cells[start_row + i, 10].Value = dt.Rows[i]["partner_id"];
				exSheet.Cells[start_row + i, 11].Value = dt.Rows[i]["partner_name"];
				exSheet.Cells[start_row + i, 12].Value = "";//dt.Rows[i]["remark"];
				//exSheet.Cells[start_row + i, 13].Value = dt.Rows[i]["accd_dr"];
				//exSheet.Cells[start_row + i, 14].Value = dt.Rows[i]["accd_cr1"];
				//exSheet.Cells[start_row + i, 15].Value = dt.Rows[i]["accd_cr2"];
				
				exSheet.Cells[start_row + i, 13].Value = "";
				exSheet.Cells[start_row + i, 14].Value = "";
				exSheet.Cells[start_row + i, 15].Value = "";
				exSheet.Cells[start_row + i, 16].Value = dt.Rows[i]["item_code"];
				exSheet.Cells[start_row + i, 17].Value = dt.Rows[i]["item_name"];
				exSheet.Cells[start_row + i, 18].Value = dt.Rows[i]["item_uom"];
				exSheet.Cells[start_row + i, 19].Value = dt.Rows[i]["qty"];
				exSheet.Cells[start_row + i, 20].Value = dt.Rows[i]["u_price"];
				exSheet.Cells[start_row + i, 21].Value = dt.Rows[i]["net_tr_amt"];
				exSheet.Cells[start_row + i, 22].Value = dt.Rows[i]["tr_rate"];
				exSheet.Cells[start_row + i, 23].Value = "";//dt.Rows[i]["net_bk_amt"];
				exSheet.Cells[start_row + i, 24].Value = dt.Rows[i]["vat_rate"];
                exSheet.Cells[start_row + i, 25].Value = dt.Rows[i]["vat_tr_amt"];
				exSheet.Cells[start_row + i, 26].Value = "";//dt.Rows[i]["vat_bk_amt"];
				exSheet.Cells[start_row + i, 27].Value = dt.Rows[i]["trans_amt"];
                exSheet.Cells[start_row + i, 28].Value = "";//dt.Rows[i]["books_amt"];
				exSheet.Cells[start_row + i, 29].Value = dt.Rows[i]["order_no"];
				exSheet.Cells[start_row + i, 30].Value = "";//dt.Rows[i]["declare_no"];
				exSheet.Cells[start_row + i, 31].Value = "";//dt.Rows[i]["remark"]; 
				exSheet.Cells[start_row + i, 32].Value = dt.Rows[i]["group_name"]; 
				exSheet.Cells[start_row + i, 33].Value = dt.Rows[i]["ecust_item_code"]; 
				
		
            }
           /* end_row = start_row + dt.Rows.Count;

            exSheet.Cells[end_row, 19].Formula = "=Sum(S" + start_row + ":S" + (end_row - 1) + ")";
            exSheet.Cells[end_row, 21].Formula = "=Sum(U" + start_row + ":U" + (end_row - 1) + ")";
			exSheet.Cells[end_row, 23].Formula = "=Sum(W" + start_row + ":W" + (end_row - 1) + ")";
            exSheet.Cells[end_row, 25].Formula = "=Sum(Y" + start_row + ":Y" + (end_row - 1) + ")";
			exSheet.Cells[end_row, 26].Formula = "=Sum(Z" + start_row + ":Z" + (end_row - 1) + ")";
            exSheet.Cells[end_row, 27].Formula = "=Sum(AA" + start_row + ":AA" + (end_row - 1) + ")";
			exSheet.Cells[end_row, 28].Formula = "=Sum(AB" + start_row + ":AB" + (end_row - 1) + ")";*/

        }
        //else
        //{
        //    Response.Write("Nodata found");
        //    Response.End();
        //}
        // end loop detail percent
        //if (File.Exists(TempFile))
        //{
        //    File.Delete(TempFile);
        //}

        //range = exSheet.Range["A1"];
        // hide row A5 
        //range.Rows.Hidden = true;

        // font bold header
        /*range = exSheet.Range["A1:AC1"];
        range.Rows[4].Font.Bold = true;*/

        exBook.SaveAs(TempFile);
		ESysLib.ExcelToPdf(TempFile);
        string pdfFilePath = TempFile.Replace(".xls", ".pdf");
        //write out to client broswer
        System.IO.FileInfo file = new System.IO.FileInfo(TempFile);
		//System.IO.FileInfo file = new System.IO.FileInfo(pdfFilePath);
        Response.Clear();
        Response.Charset = "UTF-8";
        Response.ContentEncoding = System.Text.Encoding.UTF8;
        //Add header, give a default file name for "File Download/Store as"
        Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlEncode(file.Name));
        //Add header, set file size to enable browser display download progress
        Response.AddHeader("Content-Length", file.Length.ToString());
        //Set the return string is unavailable reading for client, and must be downloaded
        Response.ContentType = "application/ms-exSheet";
		//Response.ContentType = "application/pdf";
        //Send file string to client 
		//Response.WriteFile(pdfFilePath);
        Response.WriteFile(TempFile);
        //Stop execute  
        Response.End();

    }
	
}
