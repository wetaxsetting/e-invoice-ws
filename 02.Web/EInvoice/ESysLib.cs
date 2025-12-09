using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.IO;
using System.Text;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.OracleClient;
/// <summary>
/// Summary description for ESysLib
/// </summary>
public class ESysLib : System.Web.UI.Page
{
    private static string _conString;
    public ESysLib()
    {
        //
        // TODO: Add constructor logic here
        //
    }
	public static string GetConnecting()
    {

        string dbuser = HttpContext.Current.Session["APP_DBUSER"].ToString();
        string dbpass = HttpContext.Current.Session["APP_DBPASS"].ToString();

        string conString = ConfigurationManager.AppSettings["strConn"];
        conString += "User Id=" + dbuser + ";Password=" + dbpass;

        return conString;
    }	
	
  /*  public static void ExcelToPdf(string excelPath)
    {
        try
        {
            //Convert Excel file to PDF file
            SautinSoft.ExcelToPdf x = new SautinSoft.ExcelToPdf();
			x.PageStyle.PageSize.A4();
            x.PageStyle.PageScale.Auto(); 
            x.Serial = "10003967308";

            //Set PDF format for output document
            x.OutputFormat = SautinSoft.ExcelToPdf.eOutputFormat.Pdf;

            string rootpath = System.Web.HttpContext.Current.Server.MapPath("~/");

            string pdfPath = excelPath.Replace(".xls", ".pdf");

            x.ConvertFile(excelPath, pdfPath);
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + "\n" + excelPath);
            throw (ex);
        }
    }*/
	 public static string ExecuteProcedureNotCommit(OracleConnection connection,OracleCommand command, string p_procedure, string[] paraIn)
    {

        string[] paraOut = new string[] { "" };
        DataTable dt = new DataTable();
        string result = "", procName = "", _exeStatement = "";

        try
        {
            command.Connection = connection;
            procName = "Call " + p_procedure + "( ";
            _exeStatement = procName;
            for (int j = 0; j < paraIn.Length; j++)
            {
                command.Parameters.AddWithValue(":" + j.ToString(), paraIn[j].Trim());
                command.Parameters[j].Direction = ParameterDirection.Input;
                procName += ":" + j.ToString() + ",";
                _exeStatement += "'" + paraIn[j] + "',";
            }
            procName += ":p_crt_by,";
            _exeStatement += "'" + HttpContext.Current.Session["LOGIN_ID"].ToString() + "',";

            command.Parameters.AddWithValue(":p_crt_by", HttpContext.Current.Session["LOGIN_ID"].ToString());
            command.Parameters[":p_crt_by"].Direction = ParameterDirection.Input;

            for (int j = 0; j < paraOut.Length; j++)
            {
                command.Parameters.Add(":out" + j.ToString(), OracleType.VarChar, 200);
                command.Parameters[command.Parameters.Count - 1].Direction = ParameterDirection.Output;
                procName += ":out" + j.ToString() + ",";
                _exeStatement += "'',";
            }
            procName = procName.Substring(0, procName.Length - 1) + ")";
            _exeStatement = _exeStatement.Substring(0, _exeStatement.Length - 1) + ")";

            command.CommandText = procName;
            command.ExecuteNonQuery();

            for (int j = paraOut.Length; j > 0; j--)
            {
                result += command.Parameters[command.Parameters.Count - j].Value.ToString() + "|!";
            }
            result = result.Substring(0, result.Length - 2);
            return result;
        }
        catch (Exception ex)
        {
            throw (ex);
        }
        
    }	
	public static string TableReadOpenString(string p_procedure, string[] paraIn)
	{
		
		string[] paraOut = new string[] { ""};
        string conString = GetConnecting();
		OracleConnection connection = new OracleConnection(conString);
		DataTable dt = new DataTable();
		string result = "", procName = "", _exeStatement = "";
		OracleCommand command = new OracleCommand();
		try
		{
			connection.Open();
			command.Connection = connection;
			procName = "Call " + p_procedure + "( ";
			_exeStatement = procName;
			for (int j = 0; j < paraIn.Length; j++)
			{
				command.Parameters.AddWithValue(":" + j.ToString(), paraIn[j].Trim());
				command.Parameters[j].Direction = ParameterDirection.Input;
				procName += ":" + j.ToString() + ",";
				_exeStatement += "'" + paraIn[j] + "',";
			}
			procName += ":p_crt_by,";
            _exeStatement += "'" + HttpContext.Current.Session["LOGIN_ID"].ToString() +"',";

            command.Parameters.AddWithValue(":p_crt_by", HttpContext.Current.Session["LOGIN_ID"].ToString());
			command.Parameters[":p_crt_by"].Direction = ParameterDirection.Input;

			for (int j = 0; j < paraOut.Length; j++)
			{
				command.Parameters.Add(":out" + j.ToString(), OracleType.VarChar, 200);
				command.Parameters[command.Parameters.Count - 1].Direction = ParameterDirection.Output;
				procName += ":out" + j.ToString() + ",";
				_exeStatement += "'',";
			}
			procName = procName.Substring(0, procName.Length - 1) + ")";
			_exeStatement = _exeStatement.Substring(0, _exeStatement.Length - 1) + ")";
WriteLogError(_exeStatement);
			command.CommandText = procName;
			command.ExecuteNonQuery();

			for (int j = paraOut.Length; j > 0; j--)
			{
				result += command.Parameters[command.Parameters.Count - j].Value.ToString() + "|!";
			}
			result = result.Substring(0, result.Length - 2);
			return result;
		}
		catch (Exception ex)
		{
			throw (ex);
		}
		finally
		{
			command.Dispose();
			connection.Close();
			connection.Dispose();
		}
	}
    public static void WriteLogFile(string content)
    {
        string rootpath = System.Web.HttpContext.Current.Server.MapPath("~/");

        string log_path = rootpath + "system/log/LogFile.txt";

        StreamWriter sw = new StreamWriter(log_path, true, Encoding.Unicode);
        sw.WriteLine();
        sw.WriteLine(string.Format("{0} -------------------------------------------------- ", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt")));
        sw.WriteLine(content);
        sw.WriteLine();
        sw.Close();
    }
    public static void CreateFile(string file_path, string content)
    {
        string rootpath = System.Web.HttpContext.Current.Server.MapPath("~/");

        string log_path = rootpath + file_path;
        try
        {
            StreamWriter sw = new StreamWriter(log_path, false, Encoding.Default);
            sw.WriteLine(content);
            sw.Close();
        }
        catch (Exception ex)
        {
            throw (ex);
        }
    }
    public static void EmptyFolder(string file_path)
    {
        string rootpath = System.Web.HttpContext.Current.Server.MapPath("~/");

        string log_path = rootpath + file_path;
        try
        {
            foreach (string file in Directory.GetFiles(log_path))
            {
                File.Delete(file);
            }
        }
        catch (Exception ex)
        {
            throw (ex);
        }
    }
    public static void WriteLogError(string content)
    {
        string rootpath = System.Web.HttpContext.Current.Server.MapPath("~/");
        string log_path = rootpath + "system/log/LogError.txt";
        StreamWriter sw = new StreamWriter(log_path, true, Encoding.Unicode);
        sw.WriteLine();
        sw.WriteLine(string.Format("{0} -------------------------------------------------- ", DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss tt")));
        sw.WriteLine(content);
        sw.WriteLine();
        sw.Close();
    }
    /*  public static void SetUser(string user)
      {
          string sql="", conString="";
          try
          {
              conString = ConfigurationManager.AppSettings["strConn"];
              //Update 20160815 by Fyzas
              //conString += "User Id=gasp;Password=gasp2";
              conString += "User Id=gasp;Password=" + GW_Lib.GASP_Cryptor_Util.Decrypt(ConfigurationManager.AppSettings["dbPwd"].ToString());

              OracleConnection connection = new OracleConnection(conString);
              connection.Open();
              OracleCommand command = new OracleCommand();
              sql = "select dbuser_pwd as xxx from tes_dbpassword where del_if=0 and dbuser_name='" + user.ToUpper() + "'";
              command.CommandText = sql;
              command.Connection = connection;
              OracleDataReader reader = command.ExecuteReader();

              while (reader.Read())
              {
                  string dbpass = reader["xxx"].ToString();
                  _conString = ConfigurationManager.AppSettings["strConn"];
                  _conString += "User Id=" + user + ";Password=" + dbpass;
              }
              reader.Close();
              command.Clone();
              connection.Close();
          }
          catch (Exception ex)
          {
              WriteLogError(ex.Message + "\n" + sql );
              throw (ex);
          }
      }
       public static string RenderControlHTML(string p_control_id, string p_output_control)
      {
          string control_html = "", conString = "", list_sql = "", grid_id = "";
          int col_idx;
          try
          {

              conString = ConfigurationManager.AppSettings["strConn"];
              //Update 20160815 by Fyzas
              //conString += "User Id=gasp;Password=gasp2";
              conString += "User Id=gasp;Password=" + GW_Lib.GASP_Cryptor_Util.Decrypt(ConfigurationManager.AppSettings["dbPwd"].ToString());

              OracleConnection connection = new OracleConnection(conString);
              connection.Open();
              OracleCommand command = new OracleCommand();

              command.CommandText = "Call GASP.sp_sel_write_control('" + p_control_id + "','" + p_output_control + "',:p_rtn_cur)";
              //command.CommandType = CommandType.StoredProcedure;
              command.Parameters.Add(":p_rtn_cur", OracleType.Cursor).Direction = ParameterDirection.Output;
              command.Connection = connection;
              OracleDataReader reader = command.ExecuteReader();
              if (p_output_control.ToLower() != "grid_list")
              {
                  while (reader.Read())
                  {
                      control_html = reader[0].ToString();
                  }
              }
              else
              {
                  while (reader.Read())
                  {
                      list_sql = reader[0].ToString().Trim();
                      grid_id = reader[1].ToString();
                      col_idx = int.Parse(reader[2].ToString());
                      control_html += SetGridColumnComboFormat(grid_id, col_idx, list_sql) + (char)13;
                  }
              }
              reader.Close();
              command.Dispose();
              connection.Dispose();
              connection.Close();

              //return HttpUtility.HtmlEncode(control_html);

              return control_html;
          }
          catch (Exception ex)
          {
              WriteLogError(ex.Message);
              throw (ex);
          }
      }
    public static string RenderControlHTMLEncode(string p_control_id, string p_output_control)
    {
        string control_html = "", conString = "";
        try
        {
            conString = ConfigurationManager.AppSettings["strConn"];
            //Update 20160815 by Fyzas
            //conString += "User Id=gasp;Password=gasp2";
			conString += "User Id=gasp;Password=" + GW_Lib.GASP_Cryptor_Util.Decrypt(ConfigurationManager.AppSettings["dbPwd"].ToString());

            OracleConnection connection = new OracleConnection(conString);
            connection.Open();
            OracleCommand command = new OracleCommand();

            command.CommandText = "Call GASP.sp_sel_write_control('" + p_control_id + "','" + p_output_control + "',:p_rtn_cur)";
            //command.CommandType = CommandType.StoredProcedure;
            command.Parameters.Add(":p_rtn_cur", OracleType.Cursor).Direction = ParameterDirection.Output;
            command.Connection = connection;
            OracleDataReader reader = command.ExecuteReader();

            while (reader.Read())
            {
                control_html = reader[0].ToString();
            }
            reader.Close();
            command.Dispose();
            connection.Dispose();
            connection.Close();

            return HttpUtility.HtmlEncode(control_html);

            //return control_html;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message);
            throw (ex);
        }
    }*/
    public static DataTable TableReadOpen(string sqlStatment)
    {
        OracleConnection connection = new OracleConnection(_conString);
        DataTable dt = new DataTable();
        OracleDataAdapter da = new OracleDataAdapter(sqlStatment, connection);
        try
        {
            connection.Open();
            da.Fill(dt);
            return dt;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + "\n"  + sqlStatment);
            
            throw (ex);
        }
        finally
        {
            da.Dispose();
            connection.Close();
        }
    }
    public static DataTable TableReadOpenCursor(string p_procedure, string p_parameter01, string p_parameter02)
    {
        string exeStatement = "";
        OracleConnection connection = new OracleConnection(_conString);
        DataTable dt = new DataTable();
        OracleCommand command = new OracleCommand();
        try
        {
            connection.Open();
            command.Connection = connection;
            exeStatement = "Call " + p_procedure + "(" + p_parameter01 + "," + p_parameter02 + ",:p_rtn_value)";
            command.CommandText = exeStatement;
            command.Parameters.Add(":p_rtn_value", OracleType.Cursor).Direction = ParameterDirection.Output;
            OracleDataReader reader = command.ExecuteReader();
            dt.Load(reader);
            return dt;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + " . " + exeStatement + " . " + p_procedure + " . ");
            throw (ex);
        }
        finally
        {
            connection.Close();
        }
    }

    public static DataTable TableReadOpenCursor(string p_procedure, string p_parameter)
    {
        string exeStatement = "";
        OracleConnection connection = new OracleConnection(_conString);
        DataTable dt = new DataTable();
        OracleCommand command = new OracleCommand();
        try
        {
            connection.Open();
            command.Connection = connection;
            exeStatement = "Call " + p_procedure + "(" + p_parameter + ",:p_rtn_value)";
            command.CommandText = exeStatement;
            command.Parameters.Add(":p_rtn_value", OracleType.Cursor).Direction = ParameterDirection.Output;
            OracleDataReader reader = command.ExecuteReader();
            dt.Load(reader);
            return dt;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + " . " + exeStatement + " . " + p_procedure + " . ");
            throw (ex);
        }
        finally
        {
            connection.Close();
        }
    }
    public static DataTable TableReadOpenCursor(string p_procedure)
    {
        string exeStatement = "";
        OracleConnection connection = new OracleConnection(_conString);
        DataTable dt = new DataTable();
        OracleCommand command = new OracleCommand();
        try
        {
            connection.Open();
            command.Connection = connection;
            exeStatement = "Call " + p_procedure + "(:p_rtn_value)";
            command.CommandText = exeStatement;
            command.Parameters.Add(":p_rtn_value", OracleType.Cursor).Direction = ParameterDirection.Output;
            OracleDataReader reader = command.ExecuteReader();
            dt.Load(reader);
            return dt;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + " . " + exeStatement + " . " + p_procedure + " . ");
            throw (ex);
        }
        finally
        {
            connection.Close();
        }
    }
    public static int TableReadOpenObject(string p_procedure, string p_parameter)
    {
        string exeStatement = "";
        OracleConnection connection = new OracleConnection(_conString);
        DataTable dt = new DataTable();
        int result;
        OracleCommand command = new OracleCommand();
        try
        {
            connection.Open();
            command.Connection = connection;
            exeStatement = "Call " + p_procedure + "(" + p_parameter + ",:p_return_code)";
            command.CommandText = exeStatement;
            command.Parameters.Add(":p_return_code", OracleType.Number).Direction = ParameterDirection.Output;
            OracleDataReader reader = command.ExecuteReader();
            dt.Load(reader);
            result = 1;
            return result;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + " . " + exeStatement + " . " + p_procedure + " . ");
            throw (ex);
        }
        finally
        {
            connection.Close();
        }
    }
    public static void TableUpdateOpen(string sqlStatment)
    {
        OracleConnection connection = new OracleConnection(_conString);
        OracleCommand command = new OracleCommand();
        try
        {
            command = new OracleCommand(sqlStatment, connection);
            connection.Open();
            command.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + " . " + sqlStatment + " . ");
            throw (ex);
        }
        finally
        {
            connection.Close();
        }
    }
    public static string SetGridColumnComboFormat(string id, int column, string sql)
    {
        string s = "";
        try
        {
            DataTable dt = TableReadOpen(sql);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (s.Length > 0)
                {
                    s += "|#" + dt.Rows[i][0] + ";" + dt.Rows[i][1];
                }
                else
                {
                    s += "#" + dt.Rows[i][0] + ";" + dt.Rows[i][1];
                }
            }

            s = id + ".SetComboFormat(" + column + ", '" + s + "');";
            return s;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + " . " + sql);
            throw (ex);
        }
    }

    public static string SetGridColumnComboFormatString(string id, int column, string data)
    {
        string s = "";
        try
        {
            s = id + ".SetComboFormat(" + column + ", '" + data + "');";
            return s;
        }
        catch (Exception ex)
        {
            throw (ex);
        }



    }
	public static string SetListDataPROC(string procName, string param)
    {
        string s = "DATA|";
        try
        {
            DataTable dt = TableReadOpenCursor(procName,param);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                s += dt.Rows[i][0] + "|" + dt.Rows[i][1] + "|";
            }
            if (s.Substring(s.Length - 1, 1) == "|")
            {
                s = s.Substring(0, s.Length - 1);
            }
            s = s.Replace("\"", "\"\"");
            return s;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + "\n"  + procName);
            throw (ex);
        }
    }
    public static string SetGridColumnComboFormatString2(string id, int column, string data)
    {
        string[] arr = new string[data.Length];
        string s = "";
        int i;
        char[] ch = { '|' };

        arr = data.Split(ch);
        try
        {

            for (i = 0; i < arr.Length; i++)
            {

                if (i > 0)
                    s += "|";
                s = s + "#" + arr[i] + ";" + arr[i + 1];

                i += 1;
            }

            s = SetGridColumnComboFormatString(id, column, s);
            return s;
        }
        catch (Exception ex)
        {
            throw (ex);
        }

    }

    public static string SetListDataSQL(string sql)
    {
        string s = "DATA|";
        try
        {
            DataTable dt = TableReadOpen(sql);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                s += dt.Rows[i][0] + "|" + dt.Rows[i][1] + "|";
            }
            if (s.Substring(s.Length - 1, 1) == "|")
            {
                s = s.Substring(0, s.Length - 1);
            }
            return s;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + "\n"  + sql );
            throw (ex);
        }
    }
    public static string SetListDataSQLPara(string sql, string para)
    {
        string strSQL = "";
        string[] arrPara = para.Split('|');

        for (int i = 0; i < arrPara.Length; i++)
        {
            arrPara[i] = "'" + arrPara[i].Replace(",", "','") + "'";
        }

        string s = "DATA|";
        try
        {
            strSQL = string.Format(sql, arrPara);
            DataTable dt = TableReadOpen(strSQL);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                s += dt.Rows[i][0] + "|" + dt.Rows[i][1] + "|";
            }
            if (s.Substring(s.Length - 1, 1) == "|")
            {
                s = s.Substring(0, s.Length - 1);
            }
            return s;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + "\n"  + sql );
            throw (ex);
        }
    }
    public static string SetListDataFUNC(string sql)
    {
        string s = "DATA|";
        try
        {
            DataTable dt = TableReadOpen(sql);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                s += dt.Rows[i][0];
            }
            if (s.Substring(s.Length - 1, 1) == "|")
            {
                s = s.Substring(0, s.Length - 1);
            }
            s = s.Replace("\"", "\"\"");
            return s;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + "\n"  + sql);
            throw (ex);
        }
    }
    public static string SetDataSQL(string sql)
    {
        string s = "";
        try
        {
            DataTable dt = TableReadOpen(sql);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    s += dt.Rows[i][j] + "|";
                }

            }
            if (s == "") return s;

            if (s.Substring(s.Length - 1, 1) == "|")
            {
                s = s.Substring(0, s.Length - 1);
            }
            return s;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + "\n"  + sql);
            throw (ex);
        }
    }
    public static string SetGridColumnDataSQL(string sql)
    {
        string s = "";
        try
        {
            DataTable dt = TableReadOpen(sql);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (s.Length > 0)
                {
                    s += "|#" + dt.Rows[i][0] + ";" + dt.Rows[i][1];
                }
                else
                {
                    s += "#" + dt.Rows[i][0] + ";" + dt.Rows[i][1];
                }
            }

            return s;
        }
        catch (Exception ex)
        {
            WriteLogError(ex.Message + "\n"  + sql);
            throw (ex);
        }
    }
}
