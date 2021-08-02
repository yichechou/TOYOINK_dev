using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
//using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Globalization;
using ClosedXML.Excel;
using System.Reflection;
using System.Transactions;

namespace Myclass
{
    public class MyClass
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        /// 
        /// 
        /// 月曆 fm_月曆;
        /// 
        //20210303 更新

        //****** 範例 ******
        // using Myclass;
        // public partial class Premium : Form
        // {
        //     public MyClass MyCode;
        //
        //   public Premium()
        //   {
        //     InitializeComponent();
        //     MyCode = new Myclass.MyClass();
        //      MyCode.strDbCon = MyCode.strDbConLeader;
        //      this.sqlConnection1.ConnectionString = MyCode.strDbConLeader;

        //      MyCode.strDbCon = MyCode.strDbConA01A;
        //      this.sqlConnection1.ConnectionString = MyCode.strDbConA01A;

        //     MyCode.strDbCon = "packet size=4096;user id=pwuser;password=sqlmis003;data source=192.168.128.219;persist security info=False;initial catalog=A01A;";
        //     //MyCode.strDbCon = "packet size=4096;user id=yj.chou;password=yjchou3369;data source=192.168.128.219;persist security info=False;initial catalog=Leader;";
        //   }
        // }

        //關閉程式
        //private void fm_menu_FormClosed(object sender, FormClosedEventArgs e)
        //{
        //    Environment.Exit(Environment.ExitCode);
        //}
        //****** 範例 ******
        //20210603 加入public String strDbConA01A 及txterr 錯誤訊息

        public string ERP_v4 = "192.168.128.253", AD2SERVER = "192.168.128.250", S2008X64 = "192.168.128.219", HRM = "192.168.128.219\\HRM,50502";
        public String strDbCon = "";

        //private String strDbCon = "packet size=4096;user id=yj.chou;password=yjchou3369;data source=192.168.128.219;persist security info=False;initial catalog=Leader;";
        public String strDbConA01A = "packet size=4096;user id=pwuser;password=sqlmis003;data source=192.168.128.219;persist security info=False;initial catalog=A01A;";
        public String strDbConLeader = "packet size=4096;user id=yj.chou;password=asdf0000;data source=192.168.128.219;persist security info=False;initial catalog=Leader;";
        public String prefix_table_name = "S2008X64.A01A.dbo.";
        public String str_enter = ((char)13).ToString() + ((char)10).ToString();
        public String DIRNAME = Application.StartupPath + @"\Log\";

        SqlConnection objCon, appCon, DBconn;
        SqlCommand objCmd;

        //開啟資料庫連線,可以給字串值 ERP,AD2SERVER,或是S2008X64
        public SqlConnection SqlDatabaseOpen(string connect_server = "S2008X64")
        {
            string strDbConn = "";
            //如果連線SQL Server 是HRM伺服器，則替換連線伺服器字串。(位置要在datasource替換之前)
            if (connect_server == "A01A") { strDbConn = strDbConA01A; }
            else { strDbConn = strDbCon; }
            //替換 datasource 連線值
            connect_server = (string)this.GetType().GetField(connect_server).GetValue(this);

            strDbConn = strDbConn.Replace("ServerIPAddr", connect_server);
            try
            {
                objCon = new SqlConnection(strDbConn);
                objCon.Open();
                return objCon;
            }
            catch
            {
                appCon = new SqlConnection(strDbConn);
                appCon.Open();
                return appCon;
            }
            finally
            {
            }
        }

        //取得datareader共用函式,並在關閉datareader時關閉對資料庫的連線。
        public SqlDataReader getSqlDataReader(string sqlstring, string DBServer = "S2008X64")
        {
            DBconn = SqlDatabaseOpen(DBServer);
            objCmd = new SqlCommand(sqlstring, DBconn);
            return this.objCmd.ExecuteReader(CommandBehavior.CloseConnection);
        }

        //取得dataAdapter共用函式，當dataadapter fill時資料庫連線會自動關閉。
        public SqlDataAdapter getSqlDataAdapter(string sqlstring, string DBServer = "S2008X64")
        {
            DBconn = SqlDatabaseOpen(DBServer);
            return (new SqlDataAdapter(sqlstring, DBconn));
        }

        //執行SQL語法delete、update、insert等操作
        public void sqlExecuteNonQuery(string sqlstring, string DBServer = "S2008X64")
        {
            DBconn = SqlDatabaseOpen(DBServer);
            objCmd = new SqlCommand(sqlstring, DBconn);
            objCmd.ExecuteNonQuery();
            DBconn.Close();
        }

        //刪掉sql字串中可能會有安全性影響的字元
        public string anti_sqlinjection(string Text)
        {
            //Text = Text.Replace("'", "");
            Text = Text.Replace(";", "");
            Text = Text.Replace("<", "");
            Text = Text.Replace(">", "");
            Text = Text.Replace(":", "");
            Text = Text.Replace("--", "");
            return Text;
        }

        //根据excle的路径把第一个sheel中的内容放入datatable
        public static DataTable ReadExcelToTable(string AppName,string path,string whereitem)//excel存放的路径
        {
            try
            {
                string connstring,sql;
                //连接字符串
                bool isXls = path.EndsWith(".xls");
                if (isXls)
                {
                    //Office 07以下版本 
                    connstring = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1';";
                }
                else
                {
                    //Office 07及以上版本 不能出现多余的空格 而且分号注意
                    connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";
                }

                using (OleDbConnection conn = new OleDbConnection(connstring))
                {
                    conn.Open();
                    DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字
                    string firstSheetName = sheetsName.Rows[0][2].ToString(); //得到第一个sheet的名字
                    //string sql = string.Format("SELECT * FROM [{0}] WHERE [日期] is not null", firstSheetName); //查询字符串
                    switch (AppName) 
                    {
                        case "fm_AUOCOPTC":
                            sql = string.Format("SELECT * FROM [{0}] where Number is not null and rtrim(Number) <> '' order by Number", firstSheetName); //查询字符串
                            break;
                        case "fm_AUO_NF_COPTC":
                            sql = string.Format("SELECT * FROM [{0}] where [PO NO] is not null and rtrim([PO NO]) <> '' order by [PO NO]", firstSheetName); //查询字符串
                            break;
                        default:
                            sql = string.Format("SELECT * FROM [{0}] WHERE [{1}] <> ''", firstSheetName, whereitem); //查询字符串
                            break;
                    }

                    System.Data.DataTable table = new System.Data.DataTable();

                    OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
                    DataSet set = new DataSet();
                    ada.Fill(set, "table");

                    conn.Close();
                    return set.Tables["table"];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public static DataTable ReadExcelSheetToTable(string AppName, string path,string SheetName, string whereitem)//excel存放的路径
        {
            try
            {
                string connstring, sql;
                //连接字符串
                bool isXls = path.EndsWith(".xls");
                if (isXls)
                {
                    //Office 07以下版本 
                    connstring = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1';";
                }
                else
                {
                    //Office 07及以上版本 不能出现多余的空格 而且分号注意
                    connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1';";
                }

                using (OleDbConnection conn = new OleDbConnection(connstring))
                {
                    conn.Open();
                    string str_SheetName;
                    //string KeySheetName = "%" + SheetName + "%";

                    DataTable dt_SheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字

                    DataRow[] dr_SheetsName = dt_SheetsName.Select("TABLE_NAME like '%" + SheetName + "%'");
                    if (dr_SheetsName.Length == 0) 
                    {
                        
                        return null;
                    }
                    str_SheetName = dr_SheetsName[0].ItemArray[2].ToString();
                    // 基本查詢
                    //str_SheetName = dt_sheetsName.Select("TABLE_NAME like '%"+ SheetName +"%'")[0]["TABLE_NAME"].ToString();
                    

                    DataTable dt_SheetHeaderName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, str_SheetName, null }); //得到所有sheet的名字
                    //string str_HeaderName = dr_HeaderName[0].ItemArray[3].ToString();
                    string str_HeaderName = dt_SheetHeaderName.Select("ORDINAL_POSITION = '1'")[0]["COLUMN_NAME"].ToString();

                    //firstSheetName = sheetsName.Rows[0][2].ToString(); //得到第一个sheet的名字
                    //string sql = string.Format("SELECT * FROM [{0}] WHERE [日期] is not null", firstSheetName); //查询字符串
                    switch (AppName)
                    {
                        case "fm_AUOCOPTC":
                            sql = string.Format("SELECT * FROM [{0}] where Number is not null and rtrim(Number) <> '' order by Number", SheetName); //查询字符串
                            break;
                        default:
                            sql = string.Format("SELECT * FROM [{0}] where [{1}] <> ''", str_SheetName, str_HeaderName); //查询字符串
                            break;
                    }

                    System.Data.DataTable table = new System.Data.DataTable();

                    OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
                    DataSet set = new DataSet();
                    ada.Fill(set, "table");

                    conn.Close();
                    return set.Tables["table"];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public void Sql_dgv(string sql_str, DataTable dt, DataGridView dgv)
        {
            dt.Clear();
            dgv.DataSource = null;

            try
            {
                sqlExecuteNonQuery(sql_str, "AD2SERVER");
                getSqlDataAdapter(sql_str).Fill(dt);
                dgv.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Sql_dt(string sql_str, DataTable dt)
        {
            try
            {
                sqlExecuteNonQuery(sql_str, "AD2SERVER");
                getSqlDataAdapter(sql_str).Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        //選擇日期區間判斷
        public static bool DateIntervalCheck(TextBox txt_date_s, TextBox txt_date_e)
        {
            DateTime date_s;
            DateTime date_e;
            date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            if (date_s > date_e)
            {
                MessageBox.Show("請修改日期區間", "日期格式錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            else
            {
                return true;
            }
        }


        //
        public static string[] YearMonthList(TextBox txt_date_e,int monthnum)
        {
            string[] monthlist = new string[12];
            //string str_date_y_e = txt_date_e.Text.Trim().Substring(0, 4);

            DateTime endDate,startYearMonth;
            //DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).ToShortDateString();
            endDate = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            startYearMonth = DateTime.Parse(endDate.ToString("yyyy-01-01"));
            //monthlist[0] = startDate.ToString("yyyyMM");

            monthlist[0] = startYearMonth.ToString("yyyyMM");
            for (int num = 1; num < monthnum; num++)
            {
                monthlist[num] = startYearMonth.AddMonths(num).ToString("yyyyMM");
            }

            return monthlist;
        }

        //public static string DateInterval(TextBox txt_date_s, TextBox txt_date_e)
        //{
        //    DateTime date_s;
        //    DateTime date_e;
        //    date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
        //    date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

        //    txt_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
        //    txt_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
        //}
        public void ClearControlValue(params Control[] ctols)
        {
            foreach (Control ctol in ctols)
            {
                string cType = ctols.GetType().Name;
                if (typeof(TextBox).Name == cType)
                {
                    (ctol as TextBox).Text = string.Empty;
                }
                else if (typeof(DataGrid).Name == cType)
                {
                    (ctol as DataGridView).Rows.Clear();
                }
                else if (typeof(CheckBox).Name == cType)
                {
                    (ctol as CheckBox).Checked = false;
                }
                else if (typeof(RadioButton).Name == cType)
                {
                    (ctol as RadioButton).Checked = false;
                }
                else if (typeof(Label).Name == cType)
                {
                    (ctol as Label).Text = string.Empty;
                }
            }

            ////把要清掉值的控制項ID傳入
            //ClearControlValue(TextBox1, CheckBoxList1, Label1);
        }

        //public void ClearTable(DataTable table)
        //{
        //    try
        //    {
        //        table.Clear();
        //    }
        //    catch (DataException e)
        //    {

        //    }
        //}
        public void Error_MessageBar(TextBox txterr,string str_errMessage)
        {
            txterr.Text = String.Format(@"{0}
{1}
===================", DateTime.Now.ToString(), str_errMessage);
            txterr.SelectionStart = txterr.Text.Length;
            txterr.ScrollToCaret();  //跳到遊標處 
        }

        //private void txterr_TextChanged(object sender, EventArgs e)
        //{
        //    txterr.SelectionStart = txterr.Text.Length;
        //    txterr.ScrollToCaret();  //跳到遊標處 
        //}
        public void WriteLog(string message)
        {
            string DIRNAME = Application.StartupPath + @"\Log\";
            string FILENAME = DIRNAME + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            string FormName = "",PCName;

            if (!Directory.Exists(DIRNAME))
                Directory.CreateDirectory(DIRNAME);

            if (!File.Exists(FILENAME))
            {
                // The File.Create method creates the file and opens a FileStream on the file. You neeed to close it.
                File.Create(FILENAME).Close();
            }

            PCName = Environment.MachineName;
            FormName = Form.ActiveForm.Name;

            using (StreamWriter sw = File.AppendText(FILENAME))
            {
                Log(PCName,FormName, message, sw);
            }
        }

        private void Log(string PCName, string FormName, string logMessage, TextWriter w)
        {
            w.Write("Log Entry : ");
            w.WriteLine("{0} {1}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString());
            w.WriteLine("說明:{0}-{1}-(2)", PCName, FormName, logMessage);
            w.WriteLine("-------------------------------");
        }

        public void ClosedXMLExportExcel(DataGridView dgvDataInfo, string path, string fileNameWithExtension)
        {
            //判断文件夹是否存在
            bool directoryExist = Directory.Exists(path);
            if (!directoryExist)
            {
                //创建
                Directory.CreateDirectory(path);//关联创建所有层级
            }
            //判断文件是否存在
            bool fileExist = File.Exists(path + "\\" + fileNameWithExtension);
            if (fileExist)
            {
                //删除文件
                File.Delete(path + "\\" + fileNameWithExtension);
            }

            DataTable dt = new DataTable();
            //添加列
            foreach (DataGridViewColumn column in dgvDataInfo.Columns)
            {
                dt.Columns.Add(column.HeaderText, column.ValueType);
            }
            //添加行
            foreach (DataGridViewRow row in dgvDataInfo.Rows)
            {
                dt.Rows.Add();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
                }
            }
            //保存成文件
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "導出的數據");

                //自动调整列的宽度
                wb.Worksheet(1).Columns().AdjustToContents();

                //保存文件
                wb.SaveAs(path + "\\" + fileNameWithExtension);
            }

            //打开文件
            System.Diagnostics.Process.Start(path + "\\" + fileNameWithExtension);
        }

        public static string ToFullTaiwanDate( DateTime datetime)
        {
            TaiwanCalendar taiwanCalendar = new TaiwanCalendar();

            return string.Format("民國 {0} 年 {1} 月 {2} 日",
                                    taiwanCalendar.GetYear(datetime),
                                    datetime.Month,
                                    datetime.Day);
        }

        public static string ToSimpleTaiwanDate( DateTime datetime)
        {
            TaiwanCalendar taiwanCalendar = new TaiwanCalendar();

            return string.Format("{0}/{1}/{2}",
                                    taiwanCalendar.GetYear(datetime),
                                    datetime.Month,
                                    datetime.Day);
        }

        public static string ToTaiwanDateYM(DateTime datetime)
        {
            TaiwanCalendar taiwanCalendar = new TaiwanCalendar();

            return string.Format("{0}{1}",
                                    taiwanCalendar.GetYear(datetime),
                                    datetime.Month);
        }

        public DataTable LINQResultToDataTable<T>(IEnumerable<T> Linqlist)
        {
            DataTable dt = new DataTable();


            PropertyInfo[] columns = null;

            if (Linqlist == null) return dt;

            foreach (T Record in Linqlist)
            {

                if (columns == null)
                {
                    columns = ((Type)Record.GetType()).GetProperties();
                    foreach (PropertyInfo GetProperty in columns)
                    {
                        Type IcolType = GetProperty.PropertyType;

                        if ((IcolType.IsGenericType) && (IcolType.GetGenericTypeDefinition()
                        == typeof(Nullable<>)))
                        {
                            IcolType = IcolType.GetGenericArguments()[0];
                        }

                        dt.Columns.Add(new DataColumn(GetProperty.Name, IcolType));
                    }
                }

                DataRow dr = dt.NewRow();

                foreach (PropertyInfo p in columns)
                {
                    dr[p.Name] = p.GetValue(Record, null) == null ? DBNull.Value : p.GetValue
                    (Record, null);
                }

                dt.Rows.Add(dr);
            }
            return dt;
        }
        //--------------------------------------------------------------------------
        // //保存成文件
        //using (XLWorkbook wb = new XLWorkbook())
        //{
        //    foreach (DataTable dt in ds_ERP.Tables)
        //    {
        //        wb.Worksheets.Add(dt, dt.TableName.ToString());

        //不具任何格式轉出
        //        //var ws = wb.Worksheets.Add(dt.TableName.ToString());
        //        //// The false parameter indicates that a table should not be created:
        //        //ws.FirstCell().InsertTable(dt, false);
        //        //ws.Columns().AdjustToContents();
        //    }

        ////保存文件
        //wb.SaveAs(path + "\\" + fileNameWithExtension);
        //}

        //--------------------------------------------------------------------------
        //狀態框自動置底
        //private void txterr_TextChanged(object sender, EventArgs e)
        //{
        //    txterr.SelectionStart = txterr.Text.Length;
        //    txterr.ScrollToCaret();  //跳到遊標處 
        //}
        //---------------------------------------------------------------------------



        //string sql_str_handling = String.Format(
        //        @"select * from 
        //            (select ML001 as 統制科目編號 ,
        //            (select MA003 from ACTMA where MA001 = ACTML.ML001) as 科目名稱,
        //            ML002 as 傳票日期 ,
        //            ML003+'-'+ML004+' -'+ML005 as 傳票編號,
        //            ML009 as 摘要 ,
        //            SUBSTRING( ML009 ,1, CHARINDEX (' ', ML009) -1) as 客戶名稱,
        //            (case (SUBSTRING( ML009 ,1, CHARINDEX (' ', ML009) -1)) 
        //             when 'CSOT' then '2%'
        //             when 'WCSOT' then '2%'
        //             when 'HKC-H2' then '1%'
        //             when 'CHOT' then '1%' end) as 手續費率,
        //            (case ML007 when '1' then ML008 else 0 end) as 借方金額,
        //            (case ML007 when '-1' then ML008 else 0 end)  as 貸方金額 ,
        //            (case ML007 when '1' then '借餘' when '-1' then '貸餘' end) as 借貸 
        //            from ACTML
        //            where ML006 = '623202' and ML002 >='{0}' and ML002 <= '{1}' and ML009 like '%帳款手續費%') 手續費
        //            where 手續費率 is not null
        //            order by 手續費率 desc,客戶名稱,傳票日期"
        //        , txt_date_s.Text.ToString().Trim(), txt_date_e.Text.ToString().Trim());

        /*
        private void 廠商()
        {
            //TODO: PURMA	廠商基本資料檔 加入comboBox_供應廠商 Items
            dt_COPMA = new DataTable();
            DataTable dt_交易幣別 = new DataTable();
            this.sqlDataAdapter1.SelectCommand.CommandText = "select MA001 as 客戶代號,MA001+'(' + MA002 + ')' as 客戶簡稱,MA014 as 交易幣別,連絡人 = MA005 from COPMA order by MA001";
            this.sqlDataAdapter1.Fill(dt_COPMA);
            this.cob_cust.Items.Clear();

            for (int i = 0; i < dt_COPMA.Rows.Count; i++)
            {
                this.cob_cust.Items.Add(dt_COPMA.Rows[i]["客戶簡稱"].ToString().Trim());

                if (dt_COPMA.Rows[i]["客戶代號"].ToString().Trim() == "AU-T")
                {
                    this.cob_cust.SelectedIndex = i;
                    break;
                }
            }
        }

        private void cob_cust_SelectedIndexChanged(object sender, EventArgs e)
        {
            // this.廠商();
        }
        */

        /*
        using (XLWorkbook wb_dcAll = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel))
                {
                    var ws2 = templateWB.Worksheet(3);

                    ws2.CopyTo(wb_dcAll, "折讓");
                }

                var wsheet_dcAll = wb_dcAll.Worksheet("折讓");

                //== 折讓_光阻 ================================================================
                int rows_count_dcAll = dt_dcAll.Rows.Count;
                int j = 0; int x = 0;
                string cust_dcAll = "";
                string sum_dcAll = "";

                foreach (DataRow row in dt_dcAll.Rows)
                {
                    if (cust_dcAll.ToString() != "" && row[0].ToString() != cust_dcAll.ToString())
                    {
                        wsheet_dcAll.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_dcAll.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_dcAll.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_dcAll.Cell(j + 7, 1).Value = cust_dcAll;
                        wsheet_dcAll.Cell(j + 7, 2).Value = "小計";
                        wsheet_dcAll.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcAll.Cell(j + 7, 11).FormulaA1 = "=sum(K" + (x + 7) + ":K" + (j + 6) + ")";
                        wsheet_dcAll.Cell(j + 7, 12).FormulaA1 = "=sum(L" + (x + 7) + ":L" + (j + 6) + ")";
                        wsheet_dcAll.Cell(j + 7, 13).FormulaA1 = "=sum(M" + (x + 7) + ":M" + (j + 6) + ")";
                        wsheet_dcAll.Cell(j + 7, 14).FormulaA1 = "=sum(N" + (x + 7) + ":N" + (j + 6) + ")";
                        wsheet_dcAll.Cell(j + 7, 15).FormulaA1 = "=sum(O" + (x + 7) + ":O" + (j + 6) + ")";
                        wsheet_dcAll.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcAll.Cell(j + 7, 17).FormulaA1 = "=sum(Q" + (x + 7) + ":Q" + (j + 6) + ")";

                        sum_dcAll += "N" + (j + 7) + "+";

                        x = j + 1;
                        j++;
                        //        cust = "1";
                    }

                    if (row[0].ToString() != cust_dcAll.ToString())
                    {
                        wsheet_dcAll.Cell(j + 7, 1).Value = row[0]; //客戶
                    }

                    wsheet_dcAll.Cell(j + 7, 2).Value = row[1]; //銷貨日期
                    wsheet_dcAll.Cell(j + 7, 3).Style.NumberFormat.Format = "@";
                    wsheet_dcAll.Cell(j + 7, 3).Value = row[2]; //折讓單別
                    wsheet_dcAll.Cell(j + 7, 4).Style.NumberFormat.Format = "@";
                    wsheet_dcAll.Cell(j + 7, 4).Value = row[3]; //折讓單號
                    wsheet_dcAll.Cell(j + 7, 5).Style.NumberFormat.Format = "@";
                    wsheet_dcAll.Cell(j + 7, 5).Value = row[4]; //客戶單號
                    wsheet_dcAll.Cell(j + 7, 6).Value = row[5]; //品號
                    wsheet_dcAll.Cell(j + 7, 7).Value = row[6]; //批號
                    wsheet_dcAll.Cell(j + 7, 8).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcAll.Cell(j + 7, 8).Value = row[7]; //銷貨單價
                    wsheet_dcAll.Cell(j + 7, 9).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcAll.Cell(j + 7, 9).Value = row[8]; //新單價
                    wsheet_dcAll.Cell(j + 7, 10).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcAll.Cell(j + 7, 10).Value = row[9]; //折讓差

                    wsheet_dcAll.Cell(j + 7, 11).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcAll.Cell(j + 7, 11).Value = row[10]; //銷貨數量
                    wsheet_dcAll.Cell(j + 7, 12).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcAll.Cell(j + 7, 12).Value = row[11]; //折讓金額
                    wsheet_dcAll.Cell(j + 7, 13).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcAll.Cell(j + 7, 13).Value = row[12]; //折讓稅額
                    wsheet_dcAll.Cell(j + 7, 14).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcAll.Cell(j + 7, 14).Value = row[13]; //台幣金額
                    wsheet_dcAll.Cell(j + 7, 15).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcAll.Cell(j + 7, 15).Value = row[14]; //台幣稅額
                    wsheet_dcAll.Cell(j + 7, 16).Style.NumberFormat.Format = "#,##0.0000";
                    wsheet_dcAll.Cell(j + 7, 16).Value = row[15]; //匯率
                    wsheet_dcAll.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcAll.Cell(j + 7, 17).Value = row[16]; //台幣合計
                    wsheet_dcAll.Cell(j + 7, 18).Value = row[17]; //發票號碼
                    wsheet_dcAll.Cell(j + 7, 19).Value = row[18]; //廠別

                    wsheet_dcAll.Cell(j + 7, 20).Value = row[19]; //銷退日

                    cust_dcAll = row[0].ToString().Trim();

                    if ((rows_count_dcAll - 1) == dt_dcAll.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        j++;
                        wsheet_dcAll.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_dcAll.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_dcAll.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_dcAll.Cell(j + 7, 1).Value = cust_dcAll;
                        wsheet_dcAll.Cell(j + 7, 2).Value = "小計";
                        wsheet_dcAll.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcAll.Cell(j + 7, 11).FormulaA1 = "=sum(K" + (x + 7) + ":K" + (j + 6) + ")";
                        wsheet_dcAll.Cell(j + 7, 12).FormulaA1 = "=sum(L" + (x + 7) + ":L" + (j + 6) + ")";
                        wsheet_dcAll.Cell(j + 7, 13).FormulaA1 = "=sum(M" + (x + 7) + ":M" + (j + 6) + ")";
                        wsheet_dcAll.Cell(j + 7, 14).FormulaA1 = "=sum(N" + (x + 7) + ":N" + (j + 6) + ")";
                        wsheet_dcAll.Cell(j + 7, 15).FormulaA1 = "=sum(O" + (x + 7) + ":O" + (j + 6) + ")";
                        wsheet_dcAll.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcAll.Cell(j + 7, 17).FormulaA1 = "=sum(Q" + (x + 7) + ":Q" + (j + 6) + ")";

                        sum_dcAll += "N" + (j + 7);
                        x = j + 1;

                        j++;

                        wsheet_dcAll.Cell("B4").Value = txt_date_s.Text.ToString() + " ~ " + txt_date_e.Text.ToString();
                        wsheet_dcAll.Cell("B5").Value = month;
                        wsheet_dcAll.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_dcAll.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_dcAll.Cell(j + 7, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_dcAll.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_dcAll.Cell(j + 7, 3).Value = "總計";
                        wsheet_dcAll.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcAll.Cell(j + 7, 11).FormulaA1 = "=SUMIF(B:B,\"小計\",K:K)";
                        wsheet_dcAll.Cell(j + 7, 12).FormulaA1 = "=SUMIF(B:B,\"小計\",L:L)";
                        wsheet_dcAll.Cell(j + 7, 13).FormulaA1 = "=SUMIF(B:B,\"小計\",M:M)";
                        wsheet_dcAll.Cell(j + 7, 14).FormulaA1 = "=SUMIF(B:B,\"小計\",N:N)";
                        wsheet_dcAll.Cell(j + 7, 15).FormulaA1 = "=SUMIF(B:B,\"小計\",O:O)";
                        wsheet_dcAll.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcAll.Cell(j + 7, 17).FormulaA1 = "=SUMIF(B:B,\"小計\",Q:Q)";
                    }
                    j++;
                }

                wsheet_dcAll.Position = 1;

                save_as_Trademark = txt_path.Text.ToString().Trim() + @"\\折讓" + txt_date_s.Text.ToString().Substring(0, 6) + "_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";

                wb_dcAll.SaveAs(save_as_Trademark);

                //打开文件
                System.Diagnostics.Process.Start(save_as_Trademark);
            }*/
        public void ERP_DTInputExcel(ClosedXML.Excel.IXLWorksheet wsheet, DataTable dt, int i_col, int j_row, string str_date_s, string str_date_e)
        {

            if (str_date_s == "0" || str_date_e == "0")
            {
                int i = 0;
                int j = 1;
                foreach (DataColumn Column in dt.Columns)
                {
                    wsheet.Cell(1, j).Value = dt.Columns[i].ColumnName;
                    wsheet.Cell(1, j).Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                    wsheet.Cell(1, j).Style.Font.Bold = true;
                    wsheet.Cell(1, j).Style.Fill.BackgroundColor = XLColor.FromHtml("#E0E0E0");
                    i++;
                    j++;
                }
            }
            else if (str_date_s == "RelatedVOU")
            {
                if (dt.TableName == "dt_MonthSum")
                {
                    wsheet.Cell(3, 1).Value = str_date_e;
                }
                else if (dt.TableName == "dt_QuarterSum")
                {
                    switch (str_date_e.Substring(4, 2))
                    {
                        case "03":
                            wsheet.Cell(3, 1).Value = str_date_e.Substring(0, 4) + "Q1";
                            break;
                        case "06":
                            wsheet.Cell(3, 1).Value = str_date_e.Substring(0, 4) + "Q2";
                            break;
                        case "09":
                            wsheet.Cell(3, 1).Value = str_date_e.Substring(0, 4) + "Q3";
                            break;
                        case "12":
                            wsheet.Cell(3, 1).Value = str_date_e.Substring(0, 4) + "Q4";
                            break;
                        default:
                            wsheet.Cell(3, 1).Value = str_date_e;
                            break;
                    }
                }
            }
            else
            {
                wsheet.Cell(2, 2).Value = str_date_s + "-" + str_date_e; //查詢月份區間
                wsheet.Cell(3, 2).Style.NumberFormat.Format = "@";
                wsheet.Cell(3, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期
            }
            int j_def = j_row;
            //int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                j_row = j_def;
                int row_num = 0;
                foreach (DataColumn Column in dt.Columns)
                {
                    switch (Column.ColumnName.ToString())
                    {
                        case "銷貨年月":
                        case "銷退年月":
                        case "單別":
                        case "會計別":
                        case "產品別":
                        case "產品別整":
                        case "原產品別":
                        case "商品":
                        case "存貨會計科目":
                        case "存貨科目":
                        case "年月":
                        case "科目編號":
                        case "科目層級1":
                        case "科目層級2":
                        case "科目層級3":
                        case "會計年度":
                        case "期別":
                        case "年度":
                        case "傳票年月":
                        case "傳票日期":
                        case "傳票編號":
                        case "單據日期":
                        case "單據年月":
                        case "銷貨單別":
                        case "銷貨單號":
                        case "結帳單別":
                        case "結帳單號":
                        case "結帳序號":
                        case "傳票單別":
                        case "傳票單號":
                        case "序號":
                        case "來源":
                        case "關係人代號":
                        case "品種別":
                        case "進貨單別":
                        case "進貨單號":
                        case "SI單號":
                        case "採購單別":
                        case "採購單號":
                        case "採購序號":
                        case "入庫庫別":
                        case "確認預交日":
                        case "庫存月份":
                        case "部門代號":
                            wsheet.Cell(i_col, j_row).Style.NumberFormat.Format = "@";
                            break;
                        case "總原價":
                        case "材料":
                        case "人工":
                        case "製費":
                        case "金額":
                        case "金額_材料":
                        case "金額_人工":
                        case "金額_製費":
                        case "金額_加工":
                        case "貸借金額":
                        case "本幣借方金額":
                        case "本幣貸方金額":
                        case "本幣未稅金額":
                        case "本幣金額":
                        case "未稅金額":
                        case "稅額":
                        case "總金額":
                        case "成本":
                        case "材料成本":
                        case "人工成本":
                        case "製費成本":
                        case "利潤":
                        case "罐數":
                        case "原料進貨金額":
                        case "商品進貨金額":
                        case "期末庫存額":
                        case "已沖本幣金額":
                            wsheet.Cell(i_col, j_row).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                            break;
                        case "銷貨數量":
                        case "銷貨數":
                        case "銷退數":
                        case "進貨數量":
                        case "採購數量":
                            wsheet.Cell(i_col, j_row).Style.NumberFormat.Format = "#,##0.000";
                            break;
                        case "平均單價":
                        case "利潤比率":
                        case "平均單位成本":
                        case "單位材料成本":
                        case "單位人工成本":
                        case "單位製費成本":
                        case "採購單價一":
                        case "採購金額一":
                        case "原幣進貨金額":
                        case "原幣金額":
                        case "原幣借方金額":
                        case "原幣貸方金額":
                        case "已沖原幣金額":
                            wsheet.Cell(i_col, j_row).Style.NumberFormat.Format = "#,##0.00;[RED](#,##0.00)";
                            break;
                        default:
                            break;
                    }
                    wsheet.Cell(i_col, j_row).Value = row[row_num];
                    row_num++;
                    j_row++;
                }
                i_col++;
            }
        }

    }
}
/* F22-1使用
 * CT_F22_1_SGLDT_After_Temp 自訂明細暫存表
資料行名稱	資料類型	允許NULL
銀行代號	nvarchar(10)	Checked
幣別	nvarchar(4)	Checked
銀行簡稱	nvarchar(30)	Checked
銀行帳號	nvarchar(30)	Checked
日期	nvarchar(8)	Checked
匯率	numeric(18, 4)	Checked
原幣入帳金額	numeric(18, 2)	Checked
原幣出帳金額	numeric(18, 2)	Checked
本幣入帳金額	numeric(18, 0)	Checked
本幣出帳金額	numeric(18, 0)	Checked
廠商代號	nvarchar(10)	Checked
廠商簡稱	nvarchar(30)	Checked
單據號碼	nvarchar(20)	Checked
付款銀行	nvarchar(10)	Checked
來源類型	nvarchar(20)	Checked
備註	nvarchar(255)	Checked
 * 
 */

/*  Datatable 轉 Dictionary
 *  Dictionary<string, string> dict_SPECIAL = dt_SPECIAL.AsEnumerable()
                .ToDictionary<DataRow, string, string> (
                row => row.Field<string>("ERP_NO"),
                row => row.Field<string>("SPECIAL"));

    //如果是專用料，加註淺藍色底色 比對
    if (dict_SPECIAL.ContainsKey(row[j].ToString()) == true ) 
    {
        wsheet.Cell(i + 3, j + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#00ccff");
    }

*/
