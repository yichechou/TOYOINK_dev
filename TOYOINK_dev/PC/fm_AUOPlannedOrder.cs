using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Transactions;
using System.Data.SqlClient;
using System.Data.OleDb;
using Myclass;
using System.Reflection;
using ClosedXML.Excel;
using System.Globalization;

namespace TOYOINK_dev
{
    /*//20210816 須注意 有使用[SAP].[dbo].fm_COPTC_log、[SAP].[dbo].fm_COPTD_log，同ERP欄位 前面加入 DEL_DATE
     * 
     */
    public partial class fm_AUOPlannedOrder : Form
    {
        public MyClass MyCode;

        月曆 fm_月曆;
        string str_Line;
        //存放月份
        static string[] ArrayMonthName = new string[8];
        static string[] ArrayLastReport_Month = new string[8];

        DataTable dt_Union_Excel = new DataTable(); //AUO與ERP品號對照表及公版
        
        //前回差比較報表
        DataTable dt_LastReport = new DataTable();
        DataTable dt_LastReport_C4A = new DataTable();
        DataTable dt_LastReport_C5DC6C = new DataTable();
        DataTable dt_LastReport_C5E = new DataTable();

        DataTable dt_LastReport_Month = new DataTable();

        //Excel有資料，對照表沒有
        DataTable dt_NoExist_Excel = new DataTable();
        DataTable dt_AddERP_Search = new DataTable();

        //上傳暫存表
        DataTable dt_ERPUP_C4A = new DataTable("C4A");
        DataTable dt_ERPUP_C5DC6C = new DataTable("CSDC6C");
        DataTable dt_ERPUP_C5E = new DataTable("C5E");

        //ERPUP 使用
        DataTable dt_TOERP_Temp = new DataTable();
        DataTable dt_COPTC = new DataTable("COPTC");
        DataTable dt_COPTD = new DataTable("COPTD");

        DataTable dt_tran_COPTC = new DataTable();
        DataTable dt_tran_COPTD = new DataTable();

        DataTable dt_建立者 = new DataTable();
        DataTable dt_Cust = new DataTable();
        DataTable dt_OrderLogDate = new DataTable();

        BindingSource bs_dtNoExist = new BindingSource();

        string defaultfilePath = "", save_as_AUO_Last = "";
        string path, fileNameWithExtension;
        string str_廠別 = "A01A", str_Import建立者ID = "",  str_Import建立者GP = "",  str_Import建立日期 = "";
        string  str_ERPUP建立者ID = "", str_ERPUP建立者GP = "", str_ERPUP建立日期 = "",str_ERPUP_CustID = "";
        double ft_sum採購金額 = 0, ft_sum數量合計 = 0, ft_sum包裝數量合計 = 0;
        int i, j, x, y;
        string sql_AddERP_Search_Subquery = "";
        int int_Log_Header = 3;
        string str_sql = "", str_sql_c = "", str_sql_d = "", str_sql_coptc = "", str_sql_coptd = "",sql_del_COPTCD = "", sql_del_tran_COPTCD = "";
        string str_sql_log = "", str_sql_logs = "",str_sql_ListC = "";
        string str_enter = ((char)13).ToString() + ((char)10).ToString();
        string str_sql_tran_COPTC = "", str_sql_tran_COPTD = "" ;
        int int_check_ERPUP_OK = 1;
        bool bool_changed_dgv_ERPUP_Edit= false;


        public fm_AUOPlannedOrder()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();
            //MyCode.strDbCon = MyCode.strDbConLeader;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConLeader;

            MyCode.strDbCon = MyCode.strDbConA01A;
            this.sqlConnection1.ConnectionString = MyCode.strDbConA01A;
        }

        //接收form1資料，並顯示
        public string loginName = "";
        public int CheckForm = 0;
        public void show_fmlogin_loginName(string data_loginName)
        {
            loginName = data_loginName;
        }
        public void show_fmlogin_CheckForm(int data_CheckForm)
        {
            CheckForm = data_CheckForm;
        }

        //TODO: get_sql_value() 若為文字，前面補上N'，強制為string為Unicode字符串
        private string get_sql_value(string data_type, string str_value)
        {
            string str_return = "";
            switch (data_type)
            {
                case "numeric":
                    str_return = str_value;
                    break;

                default:
                    str_return = "N'" + str_value + "'";
                    break;
            }
            return str_return;
        }

        //SQL查詢指令
        private void to_ExecuteNonQuery(string str_sql)
        {
            if (this.sqlConnection1.State == ConnectionState.Closed)
            {
                this.sqlConnection1.Open();
            }
            this.sqlCommand1.CommandText = str_sql;
            this.sqlCommand1.ExecuteNonQuery();
            this.sqlConnection1.Close();
        }
        //狀態框自動置底
        private void txterr_TextChanged(object sender, EventArgs e)
        {
            txterr.SelectionStart = txterr.Text.Length;
            txterr.ScrollToCaret();  //跳到遊標處 
        }
        //狀態框自動置底
        private void txterr_ERPUP_TextChanged(object sender, EventArgs e)
        {
            txterr_ERPUP.SelectionStart = txterr_ERPUP.Text.Length;
            txterr_ERPUP.ScrollToCaret();  //跳到遊標處 
        }
        private void fm_AUOPlannedOrder_Load(object sender, EventArgs e)
        {
            tabControl2.SelectedIndex = 1;

            //TODO:匯入ERP 可建立客戶訂單 使用者清單
            dt_建立者 = new DataTable();
            dt_Cust = new DataTable();
            string str_sql_peo = "select MF001,MF001 + MF002 as 人員,MF002,MF004 from ADMMF";
            MyCode.Sql_dt(str_sql_peo,dt_建立者);

            string str_sql_Cust = "select MA001,MA001 + MA002 as 簡稱 from COPMA order by MA001";
            MyCode.Sql_dt(str_sql_Cust, dt_Cust);

            this.cbo_Import建立者.Items.Clear();
            this.cbo_ERPUP建立者.Items.Clear();
            this.cbo_ERPUP_Cust.Items.Clear();


            string str_Import建立者 = "";
            string str_ERPUP建立者 = "";
            string str_CustID = "";
            int check = 0;


            for (int i = 0; i < dt_建立者.Rows.Count; i++)
            {
                str_Import建立者 = this.dt_建立者.Rows[i]["MF002"].ToString().Trim();
                this.cbo_Import建立者.Items.Add(dt_建立者.Rows[i]["人員"].ToString().Trim());

                str_ERPUP建立者 = this.dt_建立者.Rows[i]["MF002"].ToString().Trim();
                this.cbo_ERPUP建立者.Items.Add(dt_建立者.Rows[i]["人員"].ToString().Trim());

                if (str_Import建立者 == loginName || (loginName == "周怡甄" && str_Import建立者 == "MIS用"))
                {
                    this.cbo_Import建立者.SelectedIndex = i;
                    this.cbo_ERPUP建立者.SelectedIndex = i;

                    check = 1;
                }

            }

            for (int i = 0; i < dt_Cust.Rows.Count; i++)
            {
                str_CustID = this.dt_Cust.Rows[i]["MA001"].ToString().Trim();
                this.cbo_ERPUP_Cust.Items.Add(dt_Cust.Rows[i]["MA001"].ToString().Trim());

                if (str_CustID == "AU-TY")
                {
                    this.cbo_ERPUP_Cust.SelectedIndex = i;
                }
            }

            if (check == 0)
            {
                string str_ErrorMessage = "非採購人員不能使用";

                //MessageBox.Show(str_ErrorMessage);

                MessageBox.Show(str_ErrorMessage, "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //MyCode.Error_MessageBar(txterr,str_ErrorMessage);

                //txterr.Text += Environment.NewLine +
                //           DateTime.Now.ToString() + Environment.NewLine +
                //           "非採購人員不能使用" + Environment.NewLine +
                //           "===========";
                btn_file.Enabled = false;
                button_來源日期.Enabled = false;
                cbo_Import建立者.Enabled = false;

                //fm_login fm_login = new fm_login();

                //fm_login.Show();
                //this.Hide();
                return;
            }

            //fm_AUOPlannedOrder.WriteLog("恢復預設值");
            string filder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            txt_path.Text = filder;
            OrderLogDateList();
            textBox_來源日期.Text = DateTime.Now.ToString("yyyyMMdd");

            tctl_Import.SelectedIndex = 0;
            cbo_ERPUP_ShowM.SelectedIndex = 0;
            
            textBox_單據日期.Text = DateTime.Now.ToString("yyyyMMdd");
            lab_Nowdate.Text = DateTime.Now.ToString("yyyyMMdd");
            str_Import建立日期 = lab_Nowdate.Text.ToString();
            str_ERPUP建立日期 = lab_Nowdate.Text.ToString();
            cbo_ERPUP_M1.SelectedIndex = 0;
            cbo_ERPUP_M2.SelectedIndex = 0;
            cbo_ERPUP_M3.SelectedIndex = 0;

        }

        private void OrderLogDateList() 
        {
            dt_OrderLogDate = new DataTable();
            cbo_OrderLogDate.Items.Clear();
            cbo_LastDate.Items.Clear();
            cbo_ERPUP_OrderLogDate.Items.Clear();

            //已匯入的日期清單
            string sql_OrderLogDate = "SELECT distinct [R_DATE] FROM [CT_AUO_Order_Log] order by [R_DATE] desc";
            MyCode.Sql_dt(sql_OrderLogDate, dt_OrderLogDate);
            if (dt_OrderLogDate.Rows.Count != 0)
            {
                int c = 1;
                foreach (DataRow row_OrderLogDate in dt_OrderLogDate.Rows)
                {
                    if (c < 6)
                    {
                        cbo_OrderLogDate.Items.Add(row_OrderLogDate["R_DATE"].ToString());
                        cbo_LastDate.Items.Add(row_OrderLogDate["R_DATE"].ToString());
                        cbo_ERPUP_OrderLogDate.Items.Add(row_OrderLogDate["R_DATE"].ToString());
                    }
                    c++;
                }

                cbo_OrderLogDate.SelectedIndex = 0;
                cbo_ERPUP_OrderLogDate.SelectedIndex = 0;

                if (cbo_LastDate.Items.Count > 1)
                {
                    cbo_LastDate.SelectedIndex = 1;
                }
            }
        }


        private void CleanALL()
        {
            dgv_LastReport.DataSource = null;
            //dt_Union_Excel.Reset();
            //dt_Union_C4A.Reset();
            //dt_Union_C5DC6C.Reset();
            //dt_Union_C5E.Reset();
            //dt_Withlast_Excel.Reset();
        }
        private void btn_file_Click(object sender, EventArgs e)
        {

            openFileDialog1.Multiselect = true; // 允許選取多檔案 

            dt_Union_Excel = new DataTable(); //AUO與ERP品號對照表及公版
            //dt_Union_C4A = new DataTable();
            //dt_Union_C5DC6C = new DataTable();
            //dt_Union_C5E = new DataTable();

            //CleanALL();

            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                lst_path.Items.Clear();
                foreach (string str_OpenFilename in openFileDialog1.FileNames)
                {
                    lst_path.Items.Add(str_OpenFilename);
                }
            }
            else
            {
                return;
            }

            str_Line = cbo_ERPUP_Line.Text.ToString(); //線別名稱

            //---
            DataTable dt_temp_Excel = new DataTable();

            //與ERP品號對照表合併
            string sql_temp_Excel = "", sql_Union_Excel = "", str_ArrayMonthName = "";
            MyCode.sqlExecuteNonQuery("delete CT_AUO_Order_Temp", "AD2SERVER"); //刪除匯入暫存表

            foreach (string str_lstFilename in lst_path.Items)
            {
                dt_temp_Excel = FmReadExcelSheetToTable(str_lstFilename.ToString());

                if (dt_temp_Excel.Rows.Count != 0)
                {
                    //將數值欄位內的","刪除，並新增至暫存資料表
                    for (int i = 0; i < dt_temp_Excel.Rows.Count; i++)
                    {
                        dt_temp_Excel.Rows[i][3] = dt_temp_Excel.Rows[i][3].ToString().Replace(",", "");
                        dt_temp_Excel.Rows[i][4] = dt_temp_Excel.Rows[i][4].ToString().Replace(",", "");
                        dt_temp_Excel.Rows[i][5] = dt_temp_Excel.Rows[i][5].ToString().Replace(",", "");
                        dt_temp_Excel.Rows[i][6] = dt_temp_Excel.Rows[i][6].ToString().Replace(",", "");

                        sql_temp_Excel += String.Format(@"insert into CT_AUO_Order_Temp VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}');"
                                        , dt_temp_Excel.Rows[i][0], dt_temp_Excel.Rows[i][1], dt_temp_Excel.Rows[i][2]
                                        , dt_temp_Excel.Rows[i][3], dt_temp_Excel.Rows[i][4], dt_temp_Excel.Rows[i][5], dt_temp_Excel.Rows[i][6]) + "\r\n";
                    }
                    MyCode.sqlExecuteNonQuery(sql_temp_Excel, "AD2SERVER");
                    sql_temp_Excel = "";
                }
            }

            //將Excel匯入的檔案轉成datatable
            int j = 1;
            foreach (string str in ArrayMonthName)//存入新的DataTable
            {
                str_ArrayMonthName += "[M" + j + "] as '" + str + "',";
                j++;

                if (j == 5) 
                {
                    break;
                }
            }
            str_ArrayMonthName = str_ArrayMonthName.Substring(0, str_ArrayMonthName.Length - 1);

            sql_Union_Excel = String.Format(@"select [CT_AUO_ERPNO].[MATERIAL_TYPE],[CT_AUO_ERPNO].[ERP_NO],[CT_AUO_ERPNO].[FAB]
                                ,{0} from [CT_AUO_ERPNO]
                                left join [CT_AUO_Order_Temp] on [CT_AUO_Order_Temp].[FAB] = [CT_AUO_ERPNO].[FAB] and [CT_AUO_Order_Temp].[MATERIAL_TYPE] = [CT_AUO_ERPNO].[MATERIAL_TYPE]
                                order by [CT_AUO_ERPNO].[FAB],[ERP_NO]", str_ArrayMonthName);

            
            MyCode.Sql_dgv(sql_Union_Excel, dt_Union_Excel, dgv_ImportExcel);

            dgv_Set_ImportExcel();
            cbo_ERPUP_Line.SelectedIndex = 0;

            //Excel有資料，對照表沒有
            dt_NoExist_Excel = new DataTable();
            bs_dtNoExist = new BindingSource();
     
            string sql_NoExist_Excel = String.Format(@"SELECT [CT_AUO_Order_Temp].[MATERIAL_TYPE],[CT_AUO_Order_Temp].[FAB]
                                    FROM [CT_AUO_Order_Temp] WHERE NOT EXISTS(SELECT [FAB] ,[MATERIAL_TYPE],[ERP_NO] FROM [CT_AUO_ERPNO] 
                                    WHERE ([CT_AUO_ERPNO].[MATERIAL_TYPE]=[CT_AUO_Order_Temp].[MATERIAL_TYPE] and [CT_AUO_ERPNO].[FAB] = [CT_AUO_Order_Temp].[FAB]))");

            MyCode.Sql_dt(sql_NoExist_Excel, dt_NoExist_Excel);

            if (dt_NoExist_Excel.Rows.Count != 0)
            {
                string str_NoExist_Excel = "";

                foreach (DataRow NoExist_row in dt_NoExist_Excel.Rows)
                {
                    str_NoExist_Excel += NoExist_row["MATERIAL_TYPE"].ToString() + "-" + NoExist_row["FAB"].ToString() + ";" + Environment.NewLine;
                }
                
                DialogResult dr_NoExist_Excel = MessageBox.Show("品號對照表，缺少下述品號：" + Environment.NewLine + str_NoExist_Excel, "品號對照表缺資料，是否新增品項", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                
                MyCode.Error_MessageBar(txterr, "品號對照表，缺少下述品號：" + Environment.NewLine + str_NoExist_Excel);
                if (dr_NoExist_Excel == DialogResult.Yes)
                {
                    bs_dtNoExist.DataSource = dt_NoExist_Excel;
                    cbo_AddERP_FAB.DataBindings.Add("Text", bs_dtNoExist, "FAB", true);
                    txt_AddERP_TYPE.DataBindings.Add("Text", bs_dtNoExist, "MATERIAL_TYPE", true);
                    txt_AddERP_ERPNO.Text = "CDP-";
                    cbo_AddERP_SPECIAL.Text = "";
                    lab_AddERP_Status.Text = "共" + dt_NoExist_Excel.Rows.Count.ToString() + "筆";
                }
                //else if (dr_NoExist_Excel == DialogResult.No)
                //{
                //}
            }
            else
            {
                btn_ImportOrderLog.Enabled = true;
                btn_ImportOrderLog.BackColor = System.Drawing.Color.SeaGreen;
                btn_ImportOrderLog.ForeColor = System.Drawing.Color.White;
            }
            //dt_Withlast_Excel = dt_Union_Excel;


            //ArrayMonthName

            ////使用Linq進行查詢線別C4A
            //var Linq_C4A = from r in dt_Union_Excel.AsEnumerable()
            //               where r.Field<string>("FAB").Contains("C4A")
            //               select r;
            //dt_Union_C4A = Linq_C4A.CopyToDataTable();

            ////使用Linq進行查詢線別C5D.C6C
            //string[] Array_H11 = new string[] { "C5D", "C6C" };

            //var Linq_C5DC6C = from r in dt_Union_Excel.AsEnumerable()
            //                  where Array_H11.Contains(r.Field<string>("FAB"))
            //                  select r;
            //dt_Union_C5DC6C = Linq_C5DC6C.CopyToDataTable();

            ////使用Linq進行查詢線別C5E
            //var Linq_C5E = from r in dt_Union_Excel.AsEnumerable()
            //               where r.Field<string>("FAB").Contains("C5E")
            //               select r;
            //dt_Union_C5E = Linq_C5E.CopyToDataTable();


            //dgv_Set();
            //---

            //dgv_M1.AutoGenerateColumns = false;
            //dgv_Edit.DataSource = dt_Union_C4A;
            //dgv_Edit.Columns[0].Visible = false;
            //dgv_Edit.Columns[3].Visible = false;
            //dgv_Edit.Columns[5].Visible = false;
            //dgv_Edit.Columns[7].Visible = false;
            //dgv_Edit.Columns[9].Visible = false;
            //dgv_Edit.Columns[11].Visible = false;
            //dgv_Edit.Columns[1].Visible = true;
        }
        private void btn_AddERP_Add_Click(object sender, EventArgs e)
        {
            if (cbo_AddERP_FAB.Text.Length != 0 && txt_AddERP_TYPE.Text.Length != 0 && txt_AddERP_ERPNO.Text.Length != 0 && cbo_AddERP_SPECIAL.Text.Length !=0)
            {
                //驗證是否新增重複品號
                DataTable dt_Check_AddRepeat = new DataTable();
                DataTable dt_Check_AddRepeatERP = new DataTable();
                string sql_Check_AddRepeat = String.Format(@"select [FAB] ,[MATERIAL_TYPE],[ERP_NO] FROM [CT_AUO_ERPNO]
                                                where [FAB] = '{0}' and [MATERIAL_TYPE] = '{1}' and [ERP_NO] = '{2}'"
                                                , cbo_AddERP_FAB.Text.ToString().Trim(), txt_AddERP_TYPE.Text.ToString().Trim(), txt_AddERP_ERPNO.Text.ToString().Trim());
                string sql_Check_AddRepeatERP = String.Format(@"select [MB001] FROM [INVMB]
                                                where [MB001] = '{0}'"
                                                , txt_AddERP_ERPNO.Text.ToString().Trim());
                MyCode.Sql_dt(sql_Check_AddRepeat, dt_Check_AddRepeat);
                MyCode.Sql_dt(sql_Check_AddRepeatERP, dt_Check_AddRepeatERP);

                if (dt_Check_AddRepeat.Rows.Count != 0 || dt_Check_AddRepeatERP.Rows.Count == 0)
                {
                    if (dt_Check_AddRepeat.Rows.Count != 0)
                    {
                        string str_ErrorMessage = "錯誤-品號重複，已有該品項對照表無法新增";

                        //MessageBox.Show(str_ErrorMessage);

                        MessageBox.Show(str_ErrorMessage + Environment.NewLine, "錯誤-品號重複", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //MyCode.Error_MessageBar(txterr,str_ErrorMessage);

                        //txterr.Text += Environment.NewLine +
                        //        DateTime.Now.ToString() + Environment.NewLine +
                        //        ">> 錯誤-品號重複，已有該品項對照表無法新增" + Environment.NewLine +
                        //        "===========";
                    }
                    if (dt_Check_AddRepeatERP.Rows.Count == 0)
                    {
                        string str_ErrorMessage = "ERP品號錯誤，查無該ERP品號";

                        //MessageBox.Show(str_ErrorMessage);

                        MessageBox.Show(str_ErrorMessage + Environment.NewLine, "ERP品號錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //MyCode.Error_MessageBar(txterr,str_ErrorMessage);

                        //txterr.Text += Environment.NewLine +
                        //            DateTime.Now.ToString() + Environment.NewLine +
                        //            ">> ERP品號錯誤，查無該ERP品號" + Environment.NewLine +
                        //            "===========";
                        txt_AddERP_ERPNO.Focus();
                    }
                    
                    return;
                }



                //新增品號
                DialogResult dr_AddERP = MessageBox.Show("請再次確認資料" + Environment.NewLine 
                    + cbo_AddERP_FAB.Text.ToString().Trim() + " / " + txt_AddERP_TYPE.Text.ToString().Trim() + " / " + txt_AddERP_ERPNO.Text.ToString()
                    , "確認資料無誤", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                if (dr_AddERP == DialogResult.Yes)
                {

                    string sql_AddERP = String.Format(@"insert into [CT_AUO_ERPNO] VALUES('{0}','{1}','{2}','{3}','{4}','{5}');"
                                    , DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), str_Import建立者ID
                                    , cbo_AddERP_FAB.Text.ToString().Trim(), txt_AddERP_TYPE.Text.ToString().Trim(), txt_AddERP_ERPNO.Text.ToString().Trim(), cbo_AddERP_SPECIAL.Text.ToString());
                    
                    MyCode.sqlExecuteNonQuery(sql_AddERP, "AD2SERVER");

                    string str_ErrorMessage = "已新增品項至資料庫 " + cbo_AddERP_FAB.Text.ToString().Trim() + " / " + txt_AddERP_TYPE.Text.ToString().Trim() + " / " + txt_AddERP_ERPNO.Text.ToString();

                    MessageBox.Show(str_ErrorMessage);
                    //MyCode.Error_MessageBar(txterr,str_ErrorMessage);

                    //MessageBox.Show("已新增品項至資料庫");

                    //txterr.Text += Environment.NewLine +
                    //            DateTime.Now.ToString() + Environment.NewLine +
                    //            ">> 已新增品項至資料庫 " + cbo_AddERP_FAB.Text.ToString().Trim() + " / " + txt_AddERP_TYPE.Text.ToString().Trim() + " / " + txt_AddERP_ERPNO.Text.ToString() + Environment.NewLine +
                    //            "===========";

                    //新增後重新查詢一次
                    dt_AddERP_Search = new DataTable();
                    string sql_AddERP_Search = String.Format(@"select * from [CT_AUO_ERPNO] order by [FAB],[MATERIAL_TYPE]");
                    MyCode.Sql_dgv(sql_AddERP_Search, dt_AddERP_Search, dgv_ImportExcel);
                    dgv_Set_AddERP();

                    //Excel有資料，對照表沒有
                    dt_NoExist_Excel = new DataTable();
                    //bs_dtNoExist = new BindingSource();
                    //bs_dtNoExist.DataSource = null;
                    string sql_NoExist_Excel = String.Format(@"SELECT [CT_AUO_Order_Temp].[MATERIAL_TYPE],[CT_AUO_Order_Temp].[FAB]
                                    FROM [CT_AUO_Order_Temp] WHERE NOT EXISTS(SELECT [FAB] ,[MATERIAL_TYPE],[ERP_NO] FROM [CT_AUO_ERPNO] 
                                    WHERE ([CT_AUO_ERPNO].[MATERIAL_TYPE]=[CT_AUO_Order_Temp].[MATERIAL_TYPE] and [CT_AUO_ERPNO].[FAB] = [CT_AUO_Order_Temp].[FAB]))");

                    MyCode.Sql_dt(sql_NoExist_Excel, dt_NoExist_Excel);

                    if (dt_NoExist_Excel.Rows.Count != 0)
                    {
                        string str_NoExist_Excel = "";

                        foreach (DataRow NoExist_row in dt_NoExist_Excel.Rows)
                        {
                            str_NoExist_Excel += NoExist_row["MATERIAL_TYPE"].ToString() + "-" + NoExist_row["FAB"].ToString() + ";" + Environment.NewLine;
                        }

                        DialogResult dr_NoExist_Excel = MessageBox.Show("品號對照表，缺少下述品號：" + Environment.NewLine + str_NoExist_Excel, "品號對照表缺資料，是否新增品項", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                        if (dr_NoExist_Excel == DialogResult.Yes)
                        {
                            lab_AddERP_Status.Text = "共" + dt_NoExist_Excel.Rows.Count.ToString() + "筆";
                            bs_dtNoExist.DataSource = dt_NoExist_Excel;
                            txt_AddERP_ERPNO.Text = "CDP-";
                            //cbo_AddERP_FAB.DataBindings.Add("Text", bs_dtNoExist, "FAB", true);
                            //txt_AddERP_TYPE.DataBindings.Add("Text", bs_dtNoExist, "MATERIAL_TYPE", true);
                        }
                    }
                    else 
                    {
                        cbo_AddERP_FAB.Text = "";
                        txt_AddERP_TYPE.Text = "";
                        txt_AddERP_ERPNO.Text = "";
                        lab_AddERP_Status.Text = "";
                        cbo_AddERP_SPECIAL.Text = "";
                        dgv_ImportExcel.DataSource = null;
                        MessageBox.Show("缺少品號已新增完成" + Environment.NewLine + "請重新[選擇檔案]進行匯入", "警示",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    }
                }
                else
                {
                    return;
                }
            }
            else 
            {
                MessageBox.Show("欄位不可為空值" + Environment.NewLine , "錯誤-新增品號對照表", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_AddERP_Del_Click(object sender, EventArgs e)
        {
            sql_AddERP_Search_Subquery = "";

            if (cbo_AddERP_FAB.Text.Length != 0 && txt_AddERP_TYPE.Text.Length != 0 && txt_AddERP_ERPNO.Text.Length != 0 && cbo_AddERP_SPECIAL.Text.Length != 0)
            {
                sql_AddERP_Search_Subquery = "where";
                if (cbo_AddERP_FAB.Text.Length != 0)
                {
                    sql_AddERP_Search_Subquery += " [FAB] = '" + cbo_AddERP_FAB.Text.ToString().Trim() + "' and";
                }

                if (txt_AddERP_TYPE.Text.Length != 0)
                {
                    sql_AddERP_Search_Subquery += " [MATERIAL_TYPE] = '" + txt_AddERP_TYPE.Text.ToString().Trim() + "' and";
                }

                if (txt_AddERP_ERPNO.Text.Length != 0)
                {
                    sql_AddERP_Search_Subquery += " [ERP_NO] = '" + txt_AddERP_ERPNO.Text.ToString().Trim() + "' and";
                }
                sql_AddERP_Search_Subquery = sql_AddERP_Search_Subquery.Substring(0, sql_AddERP_Search_Subquery.Length - 4);

                string sql_AddERP_Del_Check = String.Format(@"select * from  [CT_AUO_ERPNO] {0};", sql_AddERP_Search_Subquery);
                DataTable dt_AddERP_Del_Check = new DataTable();
                MyCode.Sql_dt(sql_AddERP_Del_Check, dt_AddERP_Del_Check);
                if (dt_AddERP_Del_Check.Rows.Count == 0)
                {
                    string str_ErrorMessage = "錯誤 - 無符合資料，無法刪除" + Environment.NewLine
                             + cbo_AddERP_FAB.Text.ToString().Trim() + " / " + txt_AddERP_TYPE.Text.ToString().Trim() + " / " + txt_AddERP_ERPNO.Text.ToString() + " / " + cbo_AddERP_SPECIAL.Text.ToString();

                    MessageBox.Show("[FAB]、[MATERIAL_TYPE]、[ERP_NO]、[SPECIAL]須全部符合才能刪除，請確認資料!" + Environment.NewLine
                    + str_ErrorMessage, "錯誤-無符合資料", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    
                    //MyCode.Error_MessageBar(txterr,str_ErrorMessage);
                    //txterr.Text += Environment.NewLine +
                    //        DateTime.Now.ToString() + Environment.NewLine +
                    //        ">> 錯誤-無符合資料，無法刪除" + Environment.NewLine
                    //        + cbo_AddERP_FAB.Text.ToString().Trim() + " / " + txt_AddERP_TYPE.Text.ToString().Trim() + " / " + txt_AddERP_ERPNO.Text.ToString() 
                    //        + Environment.NewLine +
                    //        "===========";
                    return;
                }

                DialogResult dr_AddERP = MessageBox.Show("請再次確認刪除資料" +Environment.NewLine
                    + cbo_AddERP_FAB.Text.ToString().Trim() + " / " + txt_AddERP_TYPE.Text.ToString().Trim() + " / " + txt_AddERP_ERPNO.Text.ToString() + " / " + cbo_AddERP_SPECIAL.Text.ToString()
                    , "刪除資料", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
                if (dr_AddERP == DialogResult.Yes)
                {
                    string sql_AddERP_Del = String.Format(@"delete [CT_AUO_ERPNO] {0};"
                                            , sql_AddERP_Search_Subquery);

                    MyCode.sqlExecuteNonQuery(sql_AddERP_Del, "AD2SERVER");

                    string str_ErrorMessage = "資料庫已刪除品項" + Environment.NewLine
                        + cbo_AddERP_FAB.Text.ToString().Trim() + " / " + txt_AddERP_TYPE.Text.ToString().Trim() + " / " + txt_AddERP_ERPNO.Text.ToString() + " / " + cbo_AddERP_SPECIAL.Text.ToString();

                    MessageBox.Show(str_ErrorMessage);
                    //MyCode.Error_MessageBar(txterr,str_ErrorMessage);

                    //MessageBox.Show("資料庫已刪除品項" + Environment.NewLine
                    //     + cbo_AddERP_FAB.Text.ToString().Trim() + " / " + txt_AddERP_TYPE.Text.ToString().Trim() + " / " + txt_AddERP_ERPNO.Text.ToString());

                    //txterr.Text += Environment.NewLine +
                    //            DateTime.Now.ToString() + Environment.NewLine +
                    //            ">> 資料庫已刪除品項 " + cbo_AddERP_FAB.Text.ToString().Trim() + " / " + txt_AddERP_TYPE.Text.ToString().Trim() + " / " + txt_AddERP_ERPNO.Text.ToString() + Environment.NewLine +
                    //            "===========";


                    //sqlapp log
                    string str_sql_log_AddERP_Del = String.Format(
                              @"insert into develop_app_log VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')"
                              , str_Import建立者ID,  str_Import建立日期, "", "", "CT_AUO_ERPNO", "fm_AUOPlannedOrder", "刪除AUO計劃訂單_品號對照表_(" + cbo_AddERP_FAB.Text.ToString().Trim() + "-" + txt_AddERP_TYPE.Text.ToString().Trim() + "-" + txt_AddERP_ERPNO.Text.ToString() + "-" + cbo_AddERP_SPECIAL.Text.ToString() + ")", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                    MyCode.sqlExecuteNonQuery(str_sql_log_AddERP_Del, "AD2SERVER");

                    cbo_AddERP_FAB.Text = "";
                    txt_AddERP_TYPE.Text = "";
                    txt_AddERP_ERPNO.Text = "";
                    cbo_AddERP_SPECIAL.Text = "";

                    //刪除後重新查詢一次
                    dt_AddERP_Search = new DataTable();
                    string sql_AddERP_Search = String.Format(@"select * from [CT_AUO_ERPNO] order by [FAB],[MATERIAL_TYPE]");
                    MyCode.Sql_dgv(sql_AddERP_Search, dt_AddERP_Search, dgv_ImportExcel);
                    dgv_Set_AddERP();
                }
                else
                {
                    return;
                }
            }
            else
            {
                MessageBox.Show("[FAB]、[MATERIAL_TYPE]、[ERP_NO]、[SPECIAL]須全部不可為空值才能刪除，請確認資料!", "資料欄位為空值", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btn_AddERP_Search_Click(object sender, EventArgs e)
        {
            dt_AddERP_Search = new DataTable();
            sql_AddERP_Search_Subquery = "";
            AddERP_Search_Subquery();

            string sql_AddERP_Search = String.Format(@"select * from [CT_AUO_ERPNO] {0} order by [FAB],[MATERIAL_TYPE]", sql_AddERP_Search_Subquery);

            MyCode.Sql_dgv(sql_AddERP_Search, dt_AddERP_Search, dgv_ImportExcel);
            dgv_Set_AddERP();
        }

        private void AddERP_Search_Subquery()
        {
            if (cbo_AddERP_FAB.Text.Length != 0 || txt_AddERP_TYPE.Text.Length != 0 || txt_AddERP_ERPNO.Text.Length != 0)
            {
                sql_AddERP_Search_Subquery = "where";
                if (cbo_AddERP_FAB.Text.Length != 0)
                {
                    sql_AddERP_Search_Subquery += " [FAB] like '%" + cbo_AddERP_FAB.Text.ToString().Trim() + "%' and";
                }

                if (txt_AddERP_TYPE.Text.Length != 0)
                {
                    sql_AddERP_Search_Subquery += " [MATERIAL_TYPE] like '%" + txt_AddERP_TYPE.Text.ToString().Trim() + "%' and";
                }

                if (txt_AddERP_ERPNO.Text.Length != 0)
                {
                    sql_AddERP_Search_Subquery += " [ERP_NO] like '%" + txt_AddERP_ERPNO.Text.ToString().Trim() + "%' and";
                }
                sql_AddERP_Search_Subquery = sql_AddERP_Search_Subquery.Substring(0, sql_AddERP_Search_Subquery.Length - 4);
            }
            
        }
        private void dgv_ImportExcel_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int a = dgv_ImportExcel.CurrentRow.Index;

            cbo_AddERP_FAB.Text = dgv_ImportExcel.Rows[a].Cells[2].Value.ToString();
            txt_AddERP_TYPE.Text = dgv_ImportExcel.Rows[a].Cells[3].Value.ToString();
            txt_AddERP_ERPNO.Text = dgv_ImportExcel.Rows[a].Cells[4].Value.ToString();
            cbo_AddERP_SPECIAL.Text = dgv_ImportExcel.Rows[a].Cells[5].Value.ToString();
        }

        private void btn_AddERP_Clean_Click(object sender, EventArgs e)
        {
            cbo_AddERP_FAB.Text ="";
            txt_AddERP_TYPE.Text = "";
            txt_AddERP_ERPNO.Text = "";
            lab_AddERP_Status.Text = "";
            cbo_AddERP_SPECIAL.Text = "";
        }
            
        private void btn_AddERP_down_Click(object sender, EventArgs e)
        {
            bs_dtNoExist.MovePrevious();//移動到資料表的上一筆資料
            lab_AddERP_Status.Text = (int.Parse(bs_dtNoExist.Position.ToString())+ 1) + " / " + dt_NoExist_Excel.Rows.Count.ToString();
        }

        private void btn_AddERP_up_Click(object sender, EventArgs e)
        {
            bs_dtNoExist.MoveNext(); //移動到資料表的下一筆資料
            lab_AddERP_Status.Text = (int.Parse(bs_dtNoExist.Position.ToString()) + 1) + " / " + dt_NoExist_Excel.Rows.Count.ToString();
        }

        private void btn_ImportOrderLog_Click(object sender, EventArgs e)
        {
            string sql_Order_Log = "";
            try
            {
                if (dt_Union_Excel.Rows.Count != 0)
                {
                    DialogResult Result = MessageBox.Show("請再次確認資料", "確認匯入Excel", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                    if (Result == DialogResult.OK)
                    {
                        //將數值欄位內的","刪除，並新增至暫存資料表
                        for (int i = 0; i < dt_Union_Excel.Rows.Count; i++)
                        {
                            //個別月份存入ArrayMonthName[0]
                            int j = 3;
                            foreach (string str_MonthName in ArrayMonthName)//存入新的DataTable
                            {
                                sql_Order_Log += String.Format(@"insert into CT_AUO_Order_Log VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}');"
                                                , DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), str_Import建立者ID, textBox_來源日期.Text.ToString(), str_MonthName
                                                , dt_Union_Excel.Rows[i][0], dt_Union_Excel.Rows[i][1], dt_Union_Excel.Rows[i][2]
                                                , dt_Union_Excel.Rows[i][j]) + "\r\n";
                                j++;

                                if (j == 7)
                                {
                                    break;
                                }
                            }
                        }
                        MyCode.sqlExecuteNonQuery(sql_Order_Log, "AD2SERVER");
                        sql_Order_Log = "";

                        string str_ErrorMessage = "已匯入至資料庫-" + textBox_來源日期.Text.ToString();

                        MessageBox.Show(str_ErrorMessage);
                        //MyCode.Error_MessageBar(txterr,str_ErrorMessage);

                        //txterr.Text += Environment.NewLine +
                        //            DateTime.Now.ToString() + Environment.NewLine +
                        //            ">> 已匯入至資料庫-" + textBox_來源日期.Text.ToString() + Environment.NewLine +
                        //            "===========";
                        //TODO:上傳 ERP系統完成後，將單據號碼.單據日期.檔案路徑.EXCEL匯入.CFIPO畫面清除，
                        //並關閉 轉換ERP格式及上傳ERP按鈕

                        lst_path.Items.Clear();
                        dgv_ImportExcel.DataSource = null;
                        dgv_LastReport.DataSource = null;
                        dgv_ImportSearch.DataSource = null;
                        dt_Union_Excel.Clear();
                        OrderLogDateList();

                        btn_ImportOrderLog.Enabled = false;
                        btn_ImportOrderLog.BackColor = System.Drawing.SystemColors.Control;
                        btn_ImportOrderLog.ForeColor = System.Drawing.SystemColors.ControlText;

                    }
                    else if (Result == DialogResult.Cancel)
                    {
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                lab_status.Text = " 錯誤：請檢查【來源Excel格式】檔案重新上傳!!";
                txt_path.Text = "";
                string str_ErrorMessage = "【 " + ex.Message + " 】" + Environment.NewLine +
                                sql_Order_Log + Environment.NewLine +
                                "請先檢查【來源Excel格式】重新上傳 或 連絡MIS";

                MessageBox.Show(str_ErrorMessage, "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                //MyCode.Error_MessageBar(txterr, str_ErrorMessage);
                //txterr.Text += Environment.NewLine +
                //                DateTime.Now.ToString() + Environment.NewLine +
                //                "【 " + ex.Message + " 】" + Environment.NewLine +
                //                sql_Order_Log + Environment.NewLine +
                //                "請先檢查【來源Excel格式】重新上傳 或 連絡MIS" + Environment.NewLine +
                //                "===========";
                return;
            }
            //發生例外時，會自動rollback
            finally
            {

            }

        }

        private void btn_ImportSearch_Click(object sender, EventArgs e)
        {
            //取得報表月份清單
            dt_LastReport_Month = new DataTable();
            string sql_LastReport_Month = String.Format(@"select distinct [R_MONTH] from [CT_AUO_Order_Log] where [R_DATE] = '{0}'"
                                            , cbo_OrderLogDate.Text.ToString());
            MyCode.Sql_dt(sql_LastReport_Month, dt_LastReport_Month);
            int i = 0;
            foreach (DataRow RowLastReport_Month in dt_LastReport_Month.Rows)
            {
                ArrayLastReport_Month[i] = RowLastReport_Month["R_MONTH"].ToString();
                i++;
            }


            dt_LastReport = new DataTable();
            //查詢結果
            DataTable dt_ImportSeacrh = new DataTable();
            string sql_ImportSearch = String.Format(@"select [R_DATE],[MATERIAL_TYPE],[ERPNO],[FAB]
	                                  ,MAX(case [R_MONTH] when '{0}' then [AMOUNT] end) as '{0}'
                                      ,MAX(case [R_MONTH] when '{1}' then [AMOUNT] end) as '{1}'
	                                  ,MAX(case [R_MONTH] when '{2}' then [AMOUNT] end) as '{2}'
	                                  ,MAX(case [R_MONTH] when '{3}' then [AMOUNT] end) as '{3}'
                                    from [CT_AUO_Order_Log]
	                                where [R_DATE] = '{4}'
	                                group by [R_DATE] ,[MATERIAL_TYPE],[ERPNO],[FAB]
	                                order by [FAB],[MATERIAL_TYPE]"
                    , ArrayLastReport_Month[0], ArrayLastReport_Month[1], ArrayLastReport_Month[2], ArrayLastReport_Month[3], cbo_OrderLogDate.Text.ToString());

            MyCode.Sql_dgv(sql_ImportSearch, dt_ImportSeacrh, dgv_ImportSearch);
            dgv_ImportSearch.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            tctl_Import.SelectedIndex = 1;
            ////前回差報表查詢
            string sql_LastReport = String.Format(@"select NowR.[ERPNO],NowR.[FAB]
	                                    ,MAX(case NowR.[R_MONTH] when '{0}' then NowR.[AMOUNT] end) as '{0}'
	                                    ,MAX(case NowR.[R_MONTH] when '{0}' then NowR.[AMOUNT]-LastR.[AMOUNT] end) as '前回差{0}'
                                        ,MAX(case NowR.[R_MONTH] when '{1}' then NowR.[AMOUNT] end) as '{1}'
                                        ,MAX(case NowR.[R_MONTH] when '{1}' then NowR.[AMOUNT]-LastR.[AMOUNT] end) as '前回差{1}'
	                                    ,MAX(case NowR.[R_MONTH] when '{2}' then NowR.[AMOUNT] end) as '{2}'
                                        ,MAX(case NowR.[R_MONTH] when '{2}' then NowR.[AMOUNT]-LastR.[AMOUNT] end) as '前回差{2}'
	                                    ,MAX(case NowR.[R_MONTH] when '{3}' then NowR.[AMOUNT] end) as '{3}'
                                        ,MAX(case NowR.[R_MONTH] when '{3}' then NowR.[AMOUNT]-LastR.[AMOUNT] end) as '前回差{3}'
                                    from [CT_AUO_Order_Log] NowR
                                    left join [CT_AUO_Order_Log] LastR 
                                    on NowR.[FAB] = LastR.[FAB] and NowR.[MATERIAL_TYPE] = LastR.[MATERIAL_TYPE] and NowR.[R_MONTH] = LastR.[R_MONTH] and LastR.[R_DATE] = '{4}'
                                    where NowR.[R_DATE] = '{5}'
                                    group by NowR.[R_DATE] ,NowR.[MATERIAL_TYPE],NowR.[ERPNO],NowR.[FAB]
                                    order by NowR.[FAB],NowR.[MATERIAL_TYPE]"
                    , ArrayLastReport_Month[0], ArrayLastReport_Month[1], ArrayLastReport_Month[2], ArrayLastReport_Month[3], cbo_LastDate.Text.ToString(), cbo_OrderLogDate.Text.ToString());

            MyCode.Sql_dgv(sql_LastReport, dt_LastReport, dgv_LastReport);
            //dgv_LastReport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgv_Set_LastReport();

            //使用Linq進行查詢線別C4A
            var Linq_C4A = from r in dt_LastReport.AsEnumerable()
                           where r.Field<string>("FAB").Contains("C4A")
                           select r;
            dt_LastReport_C4A = Linq_C4A.CopyToDataTable();

            //使用Linq進行查詢線別C5D.C6C
            string[] Array_H11 = new string[] { "C5D", "C6C" };

            var Linq_C5DC6C = from r in dt_LastReport.AsEnumerable()
                              where Array_H11.Contains(r.Field<string>("FAB"))
                              select r;
            dt_LastReport_C5DC6C = Linq_C5DC6C.CopyToDataTable();

            //使用Linq進行查詢線別C5E
            var Linq_C5E = from r in dt_LastReport.AsEnumerable()
                           where r.Field<string>("FAB").Contains("C5E")
                           select r;
            dt_LastReport_C5E = Linq_C5E.CopyToDataTable();

            btn_Export_Excel.Enabled = true;

        }
        private void button_來源日期_Click(object sender, EventArgs e)
        {
            //TODO:單頭及單身若不為空值，表示已轉換Excel格式，需重新轉換 或 資料已上傳資料庫，需重新選擇日期
            //資料上傳資料庫後，dgv_ImportExcel會清空
            if (dt_Union_Excel.Rows.Count != 0)
            {
                DialogResult Result = MessageBox.Show("修改 來源日期 後，需重新【選擇Excel檔案】", "Excel檔案已匯入", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                if (Result == DialogResult.OK)
                {
                    lab_status.Text = "請 選擇【Excel檔案】";
                    lst_path.Items.Clear();
                    dgv_ImportExcel.DataSource = null;
                    dt_Union_Excel.Clear();

                    btn_ImportOrderLog.Enabled = false;
                    btn_ImportOrderLog.BackColor = System.Drawing.SystemColors.Control;
                    btn_ImportOrderLog.ForeColor = System.Drawing.SystemColors.ControlText;

                    //MyCode.Error_MessageBar(txterr,"修改來源日期，請重新【選擇Excel檔案】");
                    //txterr.Text += Environment.NewLine +
                    //           DateTime.Now.ToString() + Environment.NewLine +
                    //           " 修改來源日期，請重新【選擇Excel檔案】" + Environment.NewLine +
                    //           "===========";

                    this.fm_月曆 = new 月曆(this.textBox_來源日期, this.button_來源日期, "來源日期");
                }
                else if (Result == DialogResult.Cancel)
                {
                    return;
                }
            }
            else
            {
                this.fm_月曆 = new 月曆(this.textBox_來源日期, this.button_來源日期, "來源日期");

            }
        
        }

        private void textBox_來源日期_TextChanged(object sender, EventArgs e)
        {
            if (dt_OrderLogDate.Rows.Count != 0)
            {
                foreach (DataRow row_OrderLogDate in dt_OrderLogDate.Rows)
                {
                    if (textBox_來源日期.Text.ToString() == row_OrderLogDate["R_DATE"].ToString())
                    {
                        MessageBox.Show("選擇的【來源日期-" + textBox_來源日期.Text.ToString() + "】已重複匯入，請【刪除】後重新匯入；或修改【來源日期】");
                        btn_ImportOrderLog.Enabled = false;
                        btn_ImportOrderLog.BackColor = System.Drawing.SystemColors.Control;
                        btn_ImportOrderLog.ForeColor = System.Drawing.SystemColors.ControlText;

                        btn_file.Enabled = false;
                        btn_file.BackColor = System.Drawing.SystemColors.Control;
                        btn_file.ForeColor = System.Drawing.SystemColors.ControlText;

                        return;
                    }
                    else
                    {
                        btn_file.Enabled = true;
                        lab_status.Text = "請 選擇【Excel檔案】";
                    }
                }
            }

        }

        private void cbo_InputExcel_SelectedValueChanged(object sender, EventArgs e)
        {
            btn_file.Focus();
            lab_status.Text = "請 選擇【Excel檔案】";
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

        public static DataTable FmReadExcelSheetToTable(string path)//excel存放的路径
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

                    DataTable dt_SheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字
                    DataRow[] dr_SheetsName = dt_SheetsName.Select();

                    if (dr_SheetsName.Length == 0)
                    {
                        return null;
                    }
                    str_SheetName = dr_SheetsName[0].ItemArray[2].ToString();
                    // 基本查詢
                    //str_SheetName = dt_sheetsName.Select("TABLE_NAME like '%"+ SheetName +"%'")[0]["TABLE_NAME"].ToString();


                    DataTable dt_SheetHeaderName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, new object[] { null, null, str_SheetName, null }); //得到所有sheet的名字

                    string str_HeaderMonthName = "", str_ReHeaderMonthName = "";
                    //使用Linq進行查詢
                    var LinqHeaderName = from r in dt_SheetHeaderName.AsEnumerable()
                                         where r.Field<string>("COLUMN_NAME").Contains("GAP")
                                         select r.Field<string>("COLUMN_NAME");
                    int i = 0;
                    foreach (var array in LinqHeaderName)//存入新的DataTable
                    {
                        str_ReHeaderMonthName = array.Replace("/", "");
                        ArrayMonthName[i] = str_ReHeaderMonthName.Substring(str_ReHeaderMonthName.Length - 6, 6);

                        str_HeaderMonthName += "[" + array + "],";
                        i += 1;
                        if (i == 4) 
                        {
                            break;
                        }
                    }

                    str_HeaderMonthName = str_HeaderMonthName.Substring(0, str_HeaderMonthName.Length - 1);

                    string str_SelectHeader = "[MATERIAL_TYPE],[FAB],[FCST/reply]," + str_HeaderMonthName;
                    string str_filterRow = "[FCST/reply] = 'Forecast' and [FAB] in ('C4A','C5D','C5E','C6C')";

                    sql = string.Format("SELECT {0} FROM [{1}] where {2}", str_SelectHeader, str_SheetName, str_filterRow); //查询字符串


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

        private void dgv_Set_ImportExcel()
        {
            //設定 標題置中
            dgv_ImportExcel.EnableHeadersVisualStyles = false;
            dgv_ImportExcel.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //設定標題寬度
            dgv_ImportExcel.Columns[0].Width = 210;
            dgv_ImportExcel.Columns[1].Width = 210;
            dgv_ImportExcel.Columns[2].Width = 50;
            for (int i = 3; i < 7; i++)
            {
                dgv_ImportExcel.Columns[i].Width = 84;
                dgv_ImportExcel.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            //禁止標題排序
            foreach (DataGridViewColumn col in dgv_ImportExcel.Columns)
            {
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

        }

        private void dgv_Set_AddERP()
        {
            //設定 標題置中
            dgv_ImportExcel.EnableHeadersVisualStyles = false;
            dgv_ImportExcel.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            //設定標題寬度
            dgv_ImportExcel.Columns[0].Width = 140;
            dgv_ImportExcel.Columns[1].Width = 140;
            dgv_ImportExcel.Columns[2].Width = 50;
            dgv_ImportExcel.Columns[3].Width = 210;
            dgv_ImportExcel.Columns[4].Width = 210;
        }

        private void dgv_Set_LastReport()
        {
            //設定 前回差欄位標題底色及標題置中
            dgv_LastReport.EnableHeadersVisualStyles = false;
            dgv_LastReport.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            for (int i = 3; i < 10; i += 2)
            {
                dgv_LastReport.Columns[i].HeaderCell.Style.BackColor = Color.Yellow;
            }


            //設定標題寬度
            dgv_LastReport.Columns[0].Width = 210;
            //dgv_LastReport.Columns[1].Width = 200;
            dgv_LastReport.Columns[1].Width = 50;
            for (int i = 2; i < 10; i++)
            {
                dgv_LastReport.Columns[i].Width = 84;
                dgv_LastReport.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        private void dgv_Set_ERPUP_Edit()
        {
            if (dgv_ERPUP_Edit.Rows.Count == 0)
            {
                return;
            }

            //禁止變更行高
            dgv_ERPUP_Edit.AllowUserToResizeRows = false;
            
            //設定 前回差欄位標題底色及標題置中
            dgv_ERPUP_Edit.EnableHeadersVisualStyles = false;
            dgv_ERPUP_Edit.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            for (int i = 4; i < 9; i += 2)
            {
                dgv_ERPUP_Edit.Columns[i].HeaderCell.Style.BackColor = Color.NavajoWhite;
                dgv_ERPUP_Edit.Columns[i].DefaultCellStyle.BackColor = Color.NavajoWhite;
            }

            //設定標題寬度
            dgv_ERPUP_Edit.Columns[0].Width = 60;
            dgv_ERPUP_Edit.Columns[1].Width = 210;
            dgv_ERPUP_Edit.Columns[2].Width = 50;

            dgv_ERPUP_Edit.Columns[0].ReadOnly = true;
            dgv_ERPUP_Edit.Columns[1].ReadOnly = true;
            dgv_ERPUP_Edit.Columns[2].ReadOnly = true;

            //禁止變動列寬
            dgv_ERPUP_Edit.Columns[0].Resizable = DataGridViewTriState.False;
            dgv_ERPUP_Edit.Columns[1].Resizable = DataGridViewTriState.False;
            dgv_ERPUP_Edit.Columns[2].Resizable = DataGridViewTriState.False;

            dgv_ERPUP_Edit.Columns[9].Width = 300;
            dgv_ERPUP_Edit.Columns[10].Width = 300;
            dgv_ERPUP_Edit.Columns[11].Width = 300;

            //dgv_ERPUP_Edit.Columns[12].Visible = false;

            for (int i = 3; i < 9; i++)
            {
                dgv_ERPUP_Edit.Columns[i].Width = 75;
                dgv_ERPUP_Edit.Columns[i].Resizable = DataGridViewTriState.False;
                dgv_ERPUP_Edit.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            //禁止標題排序
            foreach (DataGridViewColumn col in dgv_ERPUP_Edit.Columns)
            {
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            switch (int_Log_Header)
            {
                case 3:
                case 4:
                    dgv_ERPUP_Edit.Columns[3].HeaderCell.Style.BackColor = Color.LightSkyBlue;
                    dgv_ERPUP_Edit.Columns[5].HeaderCell.Style.BackColor = SystemColors.Control;
                    dgv_ERPUP_Edit.Columns[7].HeaderCell.Style.BackColor = SystemColors.Control;
                    dgv_ERPUP_Edit.Columns[9].Visible = true;
                    dgv_ERPUP_Edit.Columns[10].Visible = false;
                    dgv_ERPUP_Edit.Columns[11].Visible = false;
                    break;
                case 5:
                case 6:
                    dgv_ERPUP_Edit.Columns[3].HeaderCell.Style.BackColor = SystemColors.Control;
                    dgv_ERPUP_Edit.Columns[5].HeaderCell.Style.BackColor = Color.LightSkyBlue;
                    dgv_ERPUP_Edit.Columns[7].HeaderCell.Style.BackColor = SystemColors.Control;
                    dgv_ERPUP_Edit.Columns[9].Visible = false;
                    dgv_ERPUP_Edit.Columns[10].Visible = true;
                    dgv_ERPUP_Edit.Columns[11].Visible = false;
                    break;
                case 7:
                case 8:
                    dgv_ERPUP_Edit.Columns[3].HeaderCell.Style.BackColor = SystemColors.Control;
                    dgv_ERPUP_Edit.Columns[5].HeaderCell.Style.BackColor = SystemColors.Control;
                    dgv_ERPUP_Edit.Columns[7].HeaderCell.Style.BackColor = Color.LightSkyBlue;
                    dgv_ERPUP_Edit.Columns[9].Visible = false;
                    dgv_ERPUP_Edit.Columns[10].Visible = false;
                    dgv_ERPUP_Edit.Columns[11].Visible = true;
                    break;
            }

        }

        private void btn_Export_Excel_Click(object sender, EventArgs e)
        {
            using (XLWorkbook wb_AUO_Last = new XLWorkbook())
            {
                wb_AUO_Last.Worksheets.Add("C4A前回差");
                wb_AUO_Last.Worksheets.Add("C5DC6C前回差");
                wb_AUO_Last.Worksheets.Add("C5E前回差");

                var wsheet_C4A = wb_AUO_Last.Worksheet("C4A前回差");
                var wsheet_C5DC6C = wb_AUO_Last.Worksheet("C5DC6C前回差");
                var wsheet_C5E = wb_AUO_Last.Worksheet("C5E前回差");

                Custom_DTInputExcel(wsheet_C4A, dt_LastReport_C4A,"C4A");
                Custom_DTInputExcel(wsheet_C5DC6C, dt_LastReport_C5DC6C,"CSDC6C");
                Custom_DTInputExcel(wsheet_C5E, dt_LastReport_C5E,"C5E");
                //自動調整欄寬
                wsheet_C4A.Columns().AdjustToContents();
                wsheet_C5DC6C.Columns().AdjustToContents();
                wsheet_C5E.Columns().AdjustToContents();

                save_as_AUO_Last = txt_path.Text.ToString().Trim() + "\\CFI_FCST_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
                wb_AUO_Last.SaveAs(save_as_AUO_Last);

                //打开文件
                System.Diagnostics.Process.Start(save_as_AUO_Last);
            }
        }

        private void btn_Path_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            //首次defaultfilePath为空，按FolderBrowserDialog默认设置（即桌面）选择
            if (defaultfilePath != "")
            {
                //设置此次默认目录为上一次选中目录
                dialog.SelectedPath = defaultfilePath;
            }

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //记录选中的目录
                defaultfilePath = dialog.SelectedPath;
                txt_path.Text = defaultfilePath;
            }
        }


        private void btn_ImportDel_Click(object sender, EventArgs e)
        {

            DialogResult Result = MessageBox.Show("刪除資料庫檔案-"+ cbo_OrderLogDate.Text.ToString()+ "，刪除後無法還原!", "刪除已匯入Excel檔案", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

            if (Result == DialogResult.OK)
            {
                string sql_OrderLog_Del = String.Format(@"delete [CT_AUO_Order_Log] where [R_DATE] = '{0}'"
                                            , cbo_OrderLogDate.Text.ToString());

                MyCode.sqlExecuteNonQuery(sql_OrderLog_Del,"AD2SERVER");

                lab_status.Text = "已刪除資料庫檔案-"+ cbo_OrderLogDate.Text.ToString();

                //MyCode.Error_MessageBar(txterr,lab_status.Text.ToString());
                //txterr.Text += Environment.NewLine +
                //            DateTime.Now.ToString() + Environment.NewLine +
                //            ">> 已刪除資料庫檔案 - "+ cbo_OrderLogDate.Text.ToString() + Environment.NewLine +
                //            "===========";
                OrderLogDateList();
                
                lst_path.Items.Clear();
                dgv_ImportExcel.DataSource = null;
                dgv_ImportSearch.DataSource = null;
                dgv_LastReport.DataSource = null;

                btn_Export_Excel.Enabled = false;
                //btn_file.Enabled = false;
                //btn_file.BackColor = System.Drawing.SystemColors.Control;
                //btn_file.ForeColor = System.Drawing.SystemColors.ControlText;

                btn_ImportOrderLog.Enabled = false;
                btn_ImportOrderLog.BackColor = System.Drawing.SystemColors.Control;
                btn_ImportOrderLog.ForeColor = System.Drawing.SystemColors.ControlText;

                //sqlapp log
                string str_sql_log_OrderLog_Del = String.Format(
                          @"insert into develop_app_log VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')"
                          , str_Import建立者ID,  str_Import建立日期, "", cbo_OrderLogDate.Text.ToString(), "CT_AUO_Order_Log", "fm_AUOPlannedOrder", "刪除AUO計劃訂單_Excel", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                MyCode.sqlExecuteNonQuery(str_sql_log_OrderLog_Del, "AD2SERVER");
            }
            else if (Result == DialogResult.Cancel)
            {
                return;
            }
            
        }

        private void cbo_OrderLogDate_SelectedValueChanged(object sender, EventArgs e)
        {
            tctl_Import.SelectedIndex = 1;
            int int_cbo_OrderLogDate ;
            int_cbo_OrderLogDate = cbo_OrderLogDate.SelectedIndex;
            if (int_cbo_OrderLogDate < (cbo_OrderLogDate.Items.Count-1))
            {
                cbo_LastDate.SelectedIndex = int_cbo_OrderLogDate + 1;
            }

        }

        private void cbo_Import建立者_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.str_Import建立者ID = this.dt_建立者.Rows[this.cbo_Import建立者.SelectedIndex]["MF001"].ToString().Trim();
            this.str_Import建立者GP = this.dt_建立者.Rows[this.cbo_Import建立者.SelectedIndex]["MF004"].ToString().Trim();
        }

        private void btn_ERPUP_Search_Click(object sender, EventArgs e)
        {
            bool_changed_dgv_ERPUP_Edit = false;
            //cbo_ERPUP_ShowM
            Clean_TOERP_Temp();
            //取得報表月份清單
            dt_LastReport_Month = new DataTable();
            string sql_LastReport_Month = String.Format(@"select distinct [R_MONTH] from [CT_AUO_Order_Log] where [R_DATE] = '{0}'"
                                            , cbo_OrderLogDate.Text.ToString());
            MyCode.Sql_dt(sql_LastReport_Month, dt_LastReport_Month);

            if (cbo_ERPUP_ShowM.Text == "前三個月")
            {
                dt_LastReport_Month.Rows.RemoveAt(3);
            }
            else
            {
                dt_LastReport_Month.Rows.RemoveAt(0);
            }

            int i = 0;
            foreach (DataRow RowLastReport_Month in dt_LastReport_Month.Rows)
            {
                ArrayLastReport_Month[i] = RowLastReport_Month["R_MONTH"].ToString();
                i++;
            }


            //查詢結果
            DataTable dt_ERPUPSeacrh = new DataTable();
            string sql_ERPUPSearch = String.Format(@"select Row_Number() OVER(Partition by [LINENAME] order by [LINENAME]) AS [SNNO],[ERPNO],[FAB]
	                                    ,MAX(case [R_MONTH] when '{0}' then [AMOUNT] else 0 end) as '{0}'
	                                    ,MAX(case [R_MONTH] when '{0}' then [AMOUNT] else 0 end) as '編輯{0}'
                                        ,MAX(case [R_MONTH] when '{1}' then [AMOUNT] else 0 end) as '{1}'
	                                    ,MAX(case [R_MONTH] when '{1}' then [AMOUNT] else 0 end) as '編輯{1}'
	                                    ,MAX(case [R_MONTH] when '{2}' then [AMOUNT] else 0 end) as '{2}'
	                                    ,'' as '編輯{2}'
	                                    ,MAX(case [R_MONTH] when '{0}' then TD020 else '' end) as '備註{0}'
                                        ,MAX(case [R_MONTH] when '{1}' then TD020 else '' end) as '備註{1}'
                                        ,MAX(case [R_MONTH] when '{2}' then TD020 else '' end) as '備註{2}'
                                from (
                                    select [R_DATE],[R_MONTH],[MATERIAL_TYPE],[ERPNO],[FAB],[AMOUNT],TD020,right(COPTD.TD002,3) as 'LINENAME' from [CT_AUO_Order_Log] 
                                    left join COPTD on TD004 = [ERPNO] and left(TD002,6) = [R_MONTH]
                                    where [R_DATE] = '{3}' and [FAB] in ('C4A') and TD001='223' and rtrim(TD002) like '%H10'
                                union all
                                    select [R_DATE],[R_MONTH],[MATERIAL_TYPE],[ERPNO],[FAB],[AMOUNT],TD020,right(COPTD.TD002,3) as 'LINENAME' from [CT_AUO_Order_Log] 
                                    left join COPTD on TD004 = [ERPNO] and left(TD002,6) = [R_MONTH]
                                    where [R_DATE] = '{3}' and [FAB] in ('C5D','C6C') and TD001='223' and rtrim(TD002) like '%H11'
                                union all
                                    select [R_DATE],[R_MONTH],[MATERIAL_TYPE],[ERPNO],[FAB],[AMOUNT],TD020,right(COPTD.TD002,3) as 'LINENAME' from [CT_AUO_Order_Log] 
                                    left join COPTD on TD004 = [ERPNO] and left(TD002,6) = [R_MONTH]
                                    where [R_DATE] = '{3}' and [FAB] in ('C5E') and TD001='223' and rtrim(TD002) like '%H14'
                                ) as a
                                where [R_DATE] = '{3}'
                                group by [R_DATE] ,[MATERIAL_TYPE],[ERPNO],[FAB],[LINENAME]
                                order by [FAB],[MATERIAL_TYPE]"
                    , ArrayLastReport_Month[0], ArrayLastReport_Month[1], ArrayLastReport_Month[2], cbo_ERPUP_OrderLogDate.Text.ToString());

            MyCode.Sql_dt(sql_ERPUPSearch, dt_ERPUPSeacrh);

            //dgv_ERPUP_Edit.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            //使用Linq進行查詢線別C4A
            var Linq_C4A = from r in dt_ERPUPSeacrh.AsEnumerable()
                           where r.Field<string>("FAB").Contains("C4A")
                           select r;
            dt_ERPUP_C4A = Linq_C4A.CopyToDataTable();

            //使用Linq進行查詢線別C5D.C6C
            string[] Array_H11 = new string[] { "C5D", "C6C" };

            var Linq_C5DC6C = from r in dt_ERPUPSeacrh.AsEnumerable()
                              where Array_H11.Contains(r.Field<string>("FAB"))
                              select r;
            dt_ERPUP_C5DC6C = Linq_C5DC6C.CopyToDataTable();

            //使用Linq進行查詢線別C5E
            var Linq_C5E = from r in dt_ERPUPSeacrh.AsEnumerable()
                           where r.Field<string>("FAB").Contains("C5E")
                           select r;
            dt_ERPUP_C5E = Linq_C5E.CopyToDataTable();

            dgv_ERPUP_Edit.DataSource = null;
            //bds_ERPUP_Edit.DataSource = dt_ERPUP_C4A;
            //dgv_ERPUP_Edit.DataSource = bds_ERPUP_Edit;


            //DataTable dt_ERPUPSeacrh = new DataTable();
            //string sql_ERPUPSearch = String.Format(@"select [ERPNO],[FAB]
            //                       ,MAX(case [R_MONTH] when '{0}' then [AMOUNT] end) as '{0}'
            //                          ,MAX(case [R_MONTH] when '{1}' then [AMOUNT] end) as '{1}'
            //                       ,MAX(case [R_MONTH] when '{2}' then [AMOUNT] end) as '{2}'
            //                        from [CT_AUO_Order_Log]
            //                     where [R_DATE] = '{3}'
            //                     group by [R_DATE] ,[MATERIAL_TYPE],[ERPNO],[FAB]
            //                     order by [FAB],[MATERIAL_TYPE]"
            //        , ArrayLastReport_Month[0], ArrayLastReport_Month[1], ArrayLastReport_Month[2], cbo_ERPUP_OrderLogDate.Text.ToString());

            //MyCode.Sql_dt(sql_ERPUPSearch, dt_ERPUPSeacrh);

            cbo_ERPUP_Line.SelectedIndex = 1;
            tctl_ERPUP.SelectedIndex = 0;
            dgv_Set_ERPUP_Edit();

            

            btn_toerp.Enabled = true;
            btn_toerp.BackColor = System.Drawing.SystemColors.Control;
            btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;

        }

        private void cbo_ERPUP_Line_SelectedIndexChanged(object sender, EventArgs e)
        {
            dgv_ERPUP_Edit.DataSource = null;
            dgv_TOERP_Temp.DataSource = null;

            switch (cbo_ERPUP_Line.Text.ToString())
            {
                case "H10-C4A":
                    //dgv_ERPUP_Edit.DataSource = dt_ERPUP_C4A;
                    bds_ERPUP_Edit.DataSource = dt_ERPUP_C4A;
                    dgv_ERPUP_Edit.DataSource = bds_ERPUP_Edit;
                    lab_ERPUP_M1.Text = ArrayLastReport_Month[0] + "H10";
                    lab_ERPUP_M2.Text = ArrayLastReport_Month[1] + "H10";
                    lab_ERPUP_M3.Text = ArrayLastReport_Month[2] + "H10";

                    break;
                case "H11-C5D.C6C":
                    //dgv_ERPUP_Edit.DataSource = dt_ERPUP_C5DC6C;
                    bds_ERPUP_Edit.DataSource = dt_ERPUP_C5DC6C;
                    dgv_ERPUP_Edit.DataSource = bds_ERPUP_Edit;
                    lab_ERPUP_M1.Text = ArrayLastReport_Month[0] + "H11";
                    lab_ERPUP_M2.Text = ArrayLastReport_Month[1] + "H11";
                    lab_ERPUP_M3.Text = ArrayLastReport_Month[2] + "H11";

                    break;
                case "H14-C5E":
                    //dgv_ERPUP_Edit.DataSource = dt_ERPUP_C5E;
                    bds_ERPUP_Edit.DataSource = dt_ERPUP_C5E;
                    dgv_ERPUP_Edit.DataSource = bds_ERPUP_Edit;
                    lab_ERPUP_M1.Text = ArrayLastReport_Month[0] + "H14";
                    lab_ERPUP_M2.Text = ArrayLastReport_Month[1] + "H14";
                    lab_ERPUP_M3.Text = ArrayLastReport_Month[2] + "H14";

                    break;
            }
            
            
            dgv_Set_ERPUP_Edit();
            tctl_ERPUP.SelectedIndex = 0;

        }
        
        private void dgv_ERPUP_Edit_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //dgv_ERPUP_Edit.Columns[e.ColumnIndex].HeaderText.ToString()
            int_Log_Header = e.ColumnIndex;
            dgv_Set_ERPUP_Edit();
   
        }

        private void dgv_ERPUP_Edit_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.Modifiers == Keys.Control) && (e.KeyCode == Keys.C) && (dgv_ERPUP_Edit.CurrentCell != null))//Ctrl+C複製指令
            {
                CopyToClipboard();
            }

            if ((e.Modifiers == Keys.Control) && (e.KeyCode == Keys.V) && (dgv_ERPUP_Edit.CurrentCell != null))//Ctrl+V貼上指令
            {
                //dgv_ERPUP_Edit[dgv_ERPUP_Edit.CurrentCell.ColumnIndex, dgv_ERPUP_Edit.CurrentCell.RowIndex].Value = Clipboard.GetText();
                PasteClipboardValue();
            }
        }

        private void dgv_ERPUP_Edit_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dgv_ERPUP_Edit.SelectedCells.Count > 0)
                dgv_ERPUP_Edit.ContextMenuStrip = contextMenuStrip1;
        }

        //=== START datagridview 右鍵複製貼上剪下 ============================
        private void tsmi_Cut_Click(object sender, EventArgs e)
        {
            //Copy to clipboard
            CopyToClipboard();

            //Clear selected cells
            foreach (DataGridViewCell dgvCell in dgv_ERPUP_Edit.SelectedCells)
                dgvCell.Value = string.Empty.Trim();
        }

        private void tsmi_Copy_Click(object sender, EventArgs e)
        {
            CopyToClipboard();
        }

        private void tsmi_Pastr_Click(object sender, EventArgs e)
        {
            //Perform paste Operation
            PasteClipboardValue();
        }
        
        private void CopyToClipboard()
        {
            //Copy to clipboard
            DataObject dataObj = dgv_ERPUP_Edit.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void PasteClipboardValue()
        {
            //Show Error if no cell is selected
            if (dgv_ERPUP_Edit.SelectedCells.Count == 0)
            {
                MessageBox.Show("Please select a cell", "Paste",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //Get the starting Cell
            DataGridViewCell startCell = GetStartCell(dgv_ERPUP_Edit);
            //Get the clipboard value in a dictionary
            Dictionary<int, Dictionary<int, string>> cbValue =
                    ClipBoardValues(Clipboard.GetText());

            int iRowIndex = startCell.RowIndex;
            foreach (int rowKey in cbValue.Keys)
            {
                int iColIndex = startCell.ColumnIndex;
                foreach (int cellKey in cbValue[rowKey].Keys)
                {
                    //Check if the index is within the limit
                    if (iColIndex <= dgv_ERPUP_Edit.Columns.Count - 1
                    && iRowIndex <= dgv_ERPUP_Edit.Rows.Count - 1)
                    {
                        DataGridViewCell cell = dgv_ERPUP_Edit[iColIndex, iRowIndex];

                        //Copy to selected cells if 'chkPasteToSelectedCells' is checked
                        //if (cell.Selected) 
                            cell.Value = cbValue[rowKey][cellKey].Trim();
                    }
                    iColIndex++;
                }
                iRowIndex++;
            }
        }
        private void tsmi_ColumnCopy_Click(object sender, EventArgs e)
        {
            //Get the starting Cell
            DataGridViewCell startCell = GetStartCell(dgv_ERPUP_Edit);
            int iColIndex = startCell.ColumnIndex;

            dgv_ERPUP_Edit.CurrentCell = dgv_ERPUP_Edit[iColIndex, 0];
            for (int int_row = 0; int_row < dgv_ERPUP_Edit.RowCount; int_row++)
            {
                dgv_ERPUP_Edit.Rows[int_row].Cells[iColIndex].Selected = true;
            }

            CopyToClipboard();
        }

        private void tsmi_ColumnPaste_Click(object sender, EventArgs e)
        {
            PasteClipboardValue();
        }
        private DataGridViewCell GetStartCell(DataGridView dgView)
        {
            //get the smallest row,column index
            if (dgView.SelectedCells.Count == 0)
                return null;

            int rowIndex = dgView.Rows.Count - 1;
            int colIndex = dgView.Columns.Count - 1;

            foreach (DataGridViewCell dgvCell in dgView.SelectedCells)
            {
                if (dgvCell.RowIndex < rowIndex)
                    rowIndex = dgvCell.RowIndex;
                if (dgvCell.ColumnIndex < colIndex)
                    colIndex = dgvCell.ColumnIndex;
            }

            return dgView[colIndex, rowIndex];
        }

        private void btn_ERPUP_TOExcel_Click(object sender, EventArgs e)
        {
            if (lab_ERPUP_M1.Text.ToString() == "")
            {
                MessageBox.Show("【單號】欄位為空值" + Environment.NewLine + "請先點選【查詢】，再執行","錯誤",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    return;
            }

            //查詢
            DataTable dt_ERP_H10 = new DataTable("C4A");
            DataTable dt_ERP_H11 = new DataTable("C5DC6C");
            DataTable dt_ERP_H14 = new DataTable("C5E");

            //lab_ERPUP_M2
            //取得單號月份
            //H10 - C4A
            string str_ListR_H10 = "'" + ArrayLastReport_Month[0] + "H10','" + ArrayLastReport_Month[1] + "H10','" + ArrayLastReport_Month[2] + "H10'";
            //H11 - C5D.C6C
            string str_ListR_H11 = "'" + ArrayLastReport_Month[0] + "H11','" + ArrayLastReport_Month[1] + "H11','" + ArrayLastReport_Month[2] + "H11'";
            //H14 - C5E
            string str_ListR_H14 = "'" + ArrayLastReport_Month[0] + "H14','" + ArrayLastReport_Month[1] + "H14','" + ArrayLastReport_Month[2] + "H14'";

            string str_sql_H10 = String.Format
                (@"select 單別
                        ,MAX(case 單號 when '{0}H10' then 單號 end) as '單號'
			            ,MAX(case 單號 when '{0}H10' then 單據日期 end) as '單據日期'
                        ,品號
			            ,MAX(case 單號 when '{0}H10' then 訂單數量 else 0 end) as '訂單數量'
			            ,MAX(case 單號 when '{1}H10' then 單號 end) as '單號'
			            ,MAX(case 單號 when '{1}H10' then 單據日期 end) as '單據日期'
			            ,品號
			            ,MAX(case 單號 when '{1}H10' then 訂單數量 else 0 end) as '訂單數量'
			            ,MAX(case 單號 when '{2}H10' then 單號 end) as '單號'
			            ,MAX(case 單號 when '{2}H10' then 單據日期 end) as '單據日期'
			            ,品號
			            ,MAX(case 單號 when '{2}H10' then 訂單數量 else 0 end) as '訂單數量'
	            from (SELECT Rtrim(TD001) as 單別,Rtrim(TD002) as 單號,Rtrim(TD004) as 品號,TD008 as 訂單數量,Rtrim(TC039) as 單據日期,Rtrim(TC004) as 客戶代號
                  FROM [A01A].[dbo].COPTD
                  left join COPTC on TC001 = TD001 and TC002 = TD002
                  where TD001 = '223' and TD002 in ({3})
                  ) as a
                group by 單別,品號", ArrayLastReport_Month[0], ArrayLastReport_Month[1], ArrayLastReport_Month[2], str_ListR_H10);
            MyCode.Sql_dt(str_sql_H10,dt_ERP_H10);

            string str_sql_H11 = String.Format
                (@"select 單別
                        ,MAX(case 單號 when '{0}H11' then 單號 end) as '單號'
			            ,MAX(case 單號 when '{0}H11' then 單據日期 end) as '單據日期'
                        ,品號
			            ,MAX(case 單號 when '{0}H11' then 訂單數量 else 0 end) as '訂單數量'
			            ,MAX(case 單號 when '{1}H11' then 單號 end) as '單號'
			            ,MAX(case 單號 when '{1}H11' then 單據日期 end) as '單據日期'
			            ,品號
			            ,MAX(case 單號 when '{1}H11' then 訂單數量 else 0 end) as '訂單數量'
			            ,MAX(case 單號 when '{2}H11' then 單號 end) as '單號'
			            ,MAX(case 單號 when '{2}H11' then 單據日期 end) as '單據日期'
			            ,品號
			            ,MAX(case 單號 when '{2}H11' then 訂單數量 else 0 end) as '訂單數量'
	            from (SELECT Rtrim(TD001) as 單別,Rtrim(TD002) as 單號,Rtrim(TD004) as 品號,TD008 as 訂單數量,Rtrim(TC039) as 單據日期,Rtrim(TC004) as 客戶代號
                  FROM [A01A].[dbo].COPTD
                  left join COPTC on TC001 = TD001 and TC002 = TD002
                  where TD001 = '223' and TD002 in ({3})
                  ) as a
                group by 單別,品號", ArrayLastReport_Month[0], ArrayLastReport_Month[1], ArrayLastReport_Month[2], str_ListR_H11);
            MyCode.Sql_dt(str_sql_H11, dt_ERP_H11);

            //        string str_sql_H11 = String.Format
            //(@"SELECT TD001 as 單別,TD002 as 單號,TD003 as 序號,TD004 as 品號,TD008 as 訂單數量,TC039 as 單據日期,TC004 as 客戶代號
            //        FROM [A01A].[dbo].COPTD
            //        left join COPTC on TC001 = TD001 and TC002 = TD002
            //        where TD001 = '223' and TD002 in ({0})
            //        order by TD002,TD003", "'202103H11', '202104H11', '202105H11'");
            //        MyCode.Sql_dt(str_sql_H11, dt_ERP_H11);

            string str_sql_H14 = String.Format
                (@"select 單別
                        ,MAX(case 單號 when '{0}H14' then 單號 end) as '單號'
			            ,MAX(case 單號 when '{0}H14' then 單據日期 end) as '單據日期'
                        ,品號
			            ,MAX(case 單號 when '{0}H14' then 訂單數量 else 0 end) as '訂單數量'
			            ,MAX(case 單號 when '{1}H14' then 單號 end) as '單號'
			            ,MAX(case 單號 when '{1}H14' then 單據日期 end) as '單據日期'
			            ,品號
			            ,MAX(case 單號 when '{1}H14' then 訂單數量 else 0 end) as '訂單數量'
			            ,MAX(case 單號 when '{2}H14' then 單號 end) as '單號'
			            ,MAX(case 單號 when '{2}H14' then 單據日期 end) as '單據日期'
			            ,品號
			            ,MAX(case 單號 when '{2}H14' then 訂單數量 else 0 end) as '訂單數量'
	            from (SELECT Rtrim(TD001) as 單別,Rtrim(TD002) as 單號,Rtrim(TD004) as 品號,TD008 as 訂單數量,Rtrim(TC039) as 單據日期,Rtrim(TC004) as 客戶代號
                  FROM [A01A].[dbo].COPTD
                  left join COPTC on TC001 = TD001 and TC002 = TD002
                  where TD001 = '223' and TD002 in ({3})
                  ) as a
                group by 單別,品號", ArrayLastReport_Month[0], ArrayLastReport_Month[1], ArrayLastReport_Month[2], str_ListR_H14);
            MyCode.Sql_dt(str_sql_H14, dt_ERP_H14);

            DataSet ds_ERP = new DataSet();
            ds_ERP.Tables.Add(dt_ERP_H10);
            ds_ERP.Tables.Add(dt_ERP_H11);
            ds_ERP.Tables.Add(dt_ERP_H14);

            ////datatable標題縮減
            //DataTable dt_ERPR_H10 = new DataTable();
            //DataTable dt_ERPR_H11 = new DataTable();
            //DataTable dt_ERPR_H14 = new DataTable();

            //DataSet ds_ERPR = new DataSet();
            //ds_ERPR.Tables.Add(dt_ERPR_H10);
            //ds_ERPR.Tables.Add(dt_ERPR_H11);
            //ds_ERPR.Tables.Add(dt_ERPR_H14);

            string str_dt_value_old = "", str_dt_value_new = "";

            foreach (DataTable dt in ds_ERP.Tables)
            {
                DataTable dt_Comparison = dt.Clone();
                foreach (DataRow dr in dt.Rows)
                {
                    //int i = 0;
                    foreach (DataColumn dc in dt.Columns)
                    {
                        if (dt.Rows.IndexOf(dr) != 0) 
                        {
                            switch (dt.Columns.IndexOf(dc)) 
                            {
                                case 0:
                                case 1:
                                case 2:
                                case 5:
                                case 6:
                                case 9:
                                case 10:
                                    str_dt_value_old = dt_Comparison.Rows[0][dt.Columns.IndexOf(dc)].ToString();
                                    str_dt_value_new = dr[dc].ToString();

                                    if (str_dt_value_old == str_dt_value_new) 
                                    {
                                        dr[dc] = "";
                                    }
                                    break;
                                default:
                                    break;
                            }
                            //i++;
                        }
                    }
                    dt_Comparison.ImportRow(dt.Rows[0]);
                }
            }


                path = txt_path.Text.ToString().Trim();
            fileNameWithExtension = "FORECAST報表用_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
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
                DialogResult myResult = MessageBox.Show
                ("是否取代原本檔案?", "檔案已存在", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (myResult == DialogResult.Yes)
                {
                    //按了是
                    //删除文件
                    File.Delete(path + "\\" + fileNameWithExtension);
                }
                else if (myResult == DialogResult.No)
                {
                    //按了否
                    return;
                }
            }

            //保存成文件
            using (XLWorkbook wb = new XLWorkbook())
            {
                foreach (DataTable dt in ds_ERP.Tables)
                {
                    wb.Worksheets.Add(dt, dt.TableName.ToString());
                    //wb.Table("C4A").Theme = XLTableTheme.None;
                    //wb.Theme = XLTableTheme.None;
                }
                
                //保存文件
                wb.SaveAs(path + "\\" + fileNameWithExtension);
            }

            //打开文件
            System.Diagnostics.Process.Start(path + "\\" + fileNameWithExtension);
        }

        private void dgv_ERPUP_Edit_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            { 
                bool_changed_dgv_ERPUP_Edit = true;
            }
        }


        private Dictionary<int, Dictionary<int, string>> ClipBoardValues(string clipboardValue)
        {
            Dictionary<int, Dictionary<int, string>>
            copyValues = new Dictionary<int, Dictionary<int, string>>();

            String[] lines = clipboardValue.Split('\n');

            for (int i = 0; i <= lines.Length - 1; i++)
            {
                copyValues[i] = new Dictionary<int, string>();
                String[] lineContent = lines[i].Split('\t');

                //if an empty cell value copied, then set the dictionary with an empty string
                //else Set value to dictionary
                if (lineContent.Length == 0)
                    copyValues[i][0] = string.Empty;
                else
                {
                    for (int j = 0; j <= lineContent.Length - 1; j++)
                        copyValues[i][j] = lineContent[j];
                }
            }
            return copyValues;
        }
        //=== END datagridview 右鍵複製貼上剪下 ============================

        // 選取整列
        DataGridViewSelectedColumnCollection selectedColumns;


        private void dgv_ERPUP_Edit_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int int_selectedColums = e.ColumnIndex;
            dgv_ERPUP_Edit.CurrentCell = dgv_ERPUP_Edit.Rows[2].Cells[int_selectedColums];
            //dgv_ERPUP_Edit.Columns[int_selectedColums].Selected = true;

            //dgv_ERPUP_Edit.SelectedColumns[0];

            //foreach (DataGridViewRow row in dgv_ERPUP_Edit.Rows)
            //{
            //    //dgv_ERPUP_Edit.CurrentCell = dgv_ERPUP_Edit.Rows[int_selectedColums].Cells[row];
            //}

            //dgv_ERPUP_Edit.CurrentCell = dgv_ERPUP_Edit.Rows[2].Cells[0];//把現行欄位移到指定的欄位  
            //dgv_ERPUP_Edit.Rows[2].Selected = true; //把該筆資料選取

            //selectedColumns = dgv_ERPUP_Edit.SelectedColumns;
            //foreach (DataGridViewColumn column in selectedColumns)
            //{
            //    column.Selected = true;
            //}
        }

        private void bds_ERPUP_Edit_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void bds_ERPUP_Edit_BindingComplete(object sender, BindingCompleteEventArgs e)
        {
            // Check if the data source has been updated, and that no error has occured.
            if (e.BindingCompleteContext ==
              BindingCompleteContext.DataSourceUpdate && e.Exception == null)
                // If not, end the current edit.
                e.Binding.BindingManagerBase.EndCurrentEdit();
        }
            int int_CheckError_TOERP_Temp = 0;

        private void tctl_ERPUP_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (tctl_ERPUP.SelectedTab.Name == "tabP_Edit" & dgv_TOERP_Temp.Rows.Count != 0 & int_check_ERPUP_OK == 0) 
            {
                DialogResult Result = MessageBox.Show("尚未上傳，將清除[整理結果]", "確認刪除[整理結果]", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                if (Result == DialogResult.OK)
                {
                    Clean_ERPUP();
                    MessageBox.Show("已刪除[整理結果]","通知",MessageBoxButtons.OK,MessageBoxIcon.Information);

                    txterr_ERPUP.Text += Environment.NewLine +
                                        DateTime.Now.ToString() + Environment.NewLine +
                                        ">> 已刪除[整理結果]" + Environment.NewLine +
                                        "===========";
                }
                else if (Result == DialogResult.Cancel)
                {
                    return;
                }
            }
        }

        private void btn_toerp_Click(object sender, EventArgs e)
        {
            int_CheckError_TOERP_Temp = 0;
            //預交日 單據日期+5個月
            //判別Datagridview 的編輯欄位是否都有值
            MyCode.sqlExecuteNonQuery("delete [CT_AUO_TOERP_Temp]", "AD2SERVER");
            int int_Column_Check = 4, int_Value_Check = 0;
            //檢查是否有空值
            for (int i = 0; i < 3; i++)
            {
                for (int int_Row_Check = 0; int_Row_Check < dgv_ERPUP_Edit.Rows.Count; int_Row_Check++)
                {
                    if (dgv_ERPUP_Edit.Rows[int_Row_Check].Cells[int_Column_Check].Value.ToString().Length == 0)
                    {
                        int_Value_Check = 1;
                    }
                }
                int_Column_Check += 2;
            }

            if (int_Value_Check == 1)
            {
                //關閉 上傳ERP按鈕
                btn_erpup.Enabled = false;
                btn_erpup.BackColor = System.Drawing.SystemColors.Control;
                btn_erpup.ForeColor = System.Drawing.SystemColors.ControlText;
                MessageBox.Show("錯誤有空值","錯誤",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }
            else
            {
                dt_tran_COPTC = new DataTable();
                dt_tran_COPTD = new DataTable();
                str_sql_tran_COPTC = "";
                str_sql_tran_COPTC = "";
                // 準備ERP上傳 字串
                str_sql_coptc = "";
                str_sql_coptd = "";
                sql_del_COPTCD = "";
                sql_del_tran_COPTCD = "";
                str_sql_logs = "";
                str_sql_ListC = "";

                switch (cbo_ERPUP_Line.Text.ToString())
                {
                    //Add_CT_AUO_TOERP_Temp
                    case "H10-C4A":
                        Check_ERPNum_Repeat_ToTemp(lab_ERPUP_M1.Text.ToString().Trim(), cbo_ERPUP_M1.Text.ToString().Trim(), dt_ERPUP_C4A, "編輯" + ArrayLastReport_Month[0].ToString().Trim(),"備註"+ ArrayLastReport_Month[0].ToString().Trim());
                        Check_ERPNum_Repeat_ToTemp(lab_ERPUP_M2.Text.ToString().Trim(), cbo_ERPUP_M2.Text.ToString().Trim(), dt_ERPUP_C4A, "編輯" + ArrayLastReport_Month[1].ToString().Trim(), "備註" + ArrayLastReport_Month[1].ToString().Trim());
                        Check_ERPNum_Repeat_ToTemp(lab_ERPUP_M3.Text.ToString().Trim(), cbo_ERPUP_M3.Text.ToString().Trim(), dt_ERPUP_C4A, "編輯" + ArrayLastReport_Month[2].ToString().Trim(), "備註" + ArrayLastReport_Month[2].ToString().Trim());
                        break;
                    case "H11-C5D.C6C":
                        Check_ERPNum_Repeat_ToTemp(lab_ERPUP_M1.Text.ToString().Trim(), cbo_ERPUP_M1.Text.ToString().Trim(), dt_ERPUP_C5DC6C, "編輯" + ArrayLastReport_Month[0].ToString().Trim(), "備註" + ArrayLastReport_Month[0].ToString().Trim());
                        Check_ERPNum_Repeat_ToTemp(lab_ERPUP_M2.Text.ToString().Trim(), cbo_ERPUP_M2.Text.ToString().Trim(), dt_ERPUP_C5DC6C, "編輯" + ArrayLastReport_Month[1].ToString().Trim(), "備註" + ArrayLastReport_Month[1].ToString().Trim());
                        Check_ERPNum_Repeat_ToTemp(lab_ERPUP_M3.Text.ToString().Trim(), cbo_ERPUP_M3.Text.ToString().Trim(), dt_ERPUP_C5DC6C, "編輯" + ArrayLastReport_Month[2].ToString().Trim(), "備註" + ArrayLastReport_Month[2].ToString().Trim());
                        break;
                    case "H14-C5E":
                        Check_ERPNum_Repeat_ToTemp(lab_ERPUP_M1.Text.ToString().Trim(), cbo_ERPUP_M1.Text.ToString().Trim(), dt_ERPUP_C5E, "編輯" + ArrayLastReport_Month[0].ToString().Trim(), "備註" + ArrayLastReport_Month[0].ToString().Trim());
                        Check_ERPNum_Repeat_ToTemp(lab_ERPUP_M2.Text.ToString().Trim(), cbo_ERPUP_M2.Text.ToString().Trim(), dt_ERPUP_C5E, "編輯" + ArrayLastReport_Month[1].ToString().Trim(), "備註" + ArrayLastReport_Month[1].ToString().Trim());
                        Check_ERPNum_Repeat_ToTemp(lab_ERPUP_M3.Text.ToString().Trim(), cbo_ERPUP_M3.Text.ToString().Trim(), dt_ERPUP_C5E, "編輯" + ArrayLastReport_Month[2].ToString().Trim(), "備註" + ArrayLastReport_Month[2].ToString().Trim());
                        break;
                }
           
                //MessageBox.Show("沒問題");
                if(int_CheckError_TOERP_Temp == 0)
                { 
                    string sql_TOERP_Temp = String.Format(@"SELECT [TTC001] as 單別,[TTC002] as 單號,[Ver] as 版次,[SNNO] as 單身序號
        ,[TTC039] as 單據日期,[ERPNO] as 品號,[FAB] as 庫別,[AMOUNT] as 數量,[ETA] as 預計交貨日,[NOTE] as 備註,[TSOURCE] as 來源碼,[TMA001] as 客戶代號
,[TCOMPANY] as 公司別,[TCREATOR] as 建立者,[TUSR_GROUP] as 建立群組,[TCREATE_DATE] as 建立日期
         FROM [CT_AUO_TOERP_Temp] order by [TTC002] ,[SNNO]");
                    MyCode.Sql_dgv(sql_TOERP_Temp, dt_TOERP_Temp, dgv_TOERP_Temp);

                    //列出轉換成ERP格式 單頭及單身
                    dgv_COPTD.DataSource = dt_tran_COPTD;
                    dgv_COPTC.DataSource = dt_tran_COPTC;

                    lab_status.Text = " ERP格式轉換完成";
                    tctl_ERPUP.SelectedIndex = 3;
                    tctl_ERPUP.SelectedIndex = 2;

                    txterr_ERPUP.Text += Environment.NewLine +
                                    DateTime.Now.ToString() + Environment.NewLine +
                                   ">> ERP格式轉換完成" + Environment.NewLine +
                                   "===========";
                    tctl_ERPUP.SelectedIndex = 1;

                    //確認資料轉換成 ERP格式後，開啟 上傳ERP按鈕
                    btn_erpup.Enabled = true;
                    btn_erpup.BackColor = System.Drawing.Color.SteelBlue;
                    btn_erpup.ForeColor = System.Drawing.Color.White;
                }
            }

        }
        
        private void Check_ERPNum_Repeat_ToTemp( string str_ERPNum_TC002, string str_ERPNum_Status, DataTable dt_ToTemp,string str_AMOUNT,string str_NOTE) 
        {
            if (str_ERPNum_Status == "無") 
            {
                return;
            }

            DateTime str_TC_Date = DateTime.ParseExact(textBox_單據日期.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            string str_ETA = DateTime.Parse(str_TC_Date.ToString("yyyy-MM-01")).AddMonths(5).ToString("yyyyMMdd");

            string sql_ToTemp = "";

            foreach (DataRow row in dt_ToTemp.Rows)
            {
                sql_ToTemp += String.Format(@"INSERT CT_AUO_TOERP_Temp
(TTC001, TTC002, Ver, SNNO, TTC039, ERPNO, FAB, AMOUNT, ETA, NOTE, TSOURCE,TMA001, TCOMPANY, TCREATOR, TUSR_GROUP, TCREATE_DATE)
VALUES('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}','{11}','{12}', '{13}', '{14}', '{15}');",
lab_TC001.Text.ToString().Trim(), str_ERPNum_TC002, "0000", row["SNNO"].ToString().Trim().PadLeft(4, '0'),
textBox_單據日期.Text.ToString().Trim(), row["ERPNO"].ToString().Trim(), row["FAB"].ToString().Trim(), row[str_AMOUNT].ToString().Trim(),
str_ETA, row[str_NOTE].ToString().Trim(), "9", str_ERPUP_CustID, str_廠別 , str_ERPUP建立者ID, str_ERPUP建立者GP, str_ERPUP建立日期) +"\r\n";
            }
            //新增至CT_AUO_TOERP_Temp
            MyCode.sqlExecuteNonQuery(sql_ToTemp, "AD2SERVER");

            //單身 ERP格式,TD021 確認碼
            DataTable dt_單身 = new DataTable();
            string str_sql_td = String.Format(@"select [TTC001] as TD001, [TTC002] as TD002
                                ,[SNNO] as TD003,INVMB.MB001 as TD004,INVMB.MB002 as TD005,INVMB.MB003 as TD006
                                ,INVMB.MB017 as TD007,[AMOUNT] as TD008,'0' as TD009,INVMB.MB004 as TD010
                                ,'0' as TD011,'0' as TD012,[ETA] as TD013
                                ,COPMG.MG003 as TD014,'' as TD015,'N' as TD016,'' as TD017,'' as TD018,'' as TD019
                                ,[NOTE] as TD020,'Y' as TD021,'0' as TD022,'' as TD023,'0' as TD024,'0' as TD025,'1' as TD026
                                ,'' as TD027,'' as TD028,'' as TD029,'0' as TD030,'0' as TD031,'0' as TD032
                                ,'0' as TD033,'0' as TD034,'0' as TD035,INVMB.MB090 as TD036,'' as TD037,'' as TD038
                                ,'' as TD039,'' as TD040,'' as TD041,'0' as TD042,'' as TD043,'' as TD044,[TSOURCE] as TD045,'' as TD046
                                ,[ETA] as TD047,[ETA] as TD048,'1' as TD049,'0' as TD050
                                ,'0' as TD051,'0' as TD052,'0' as TD053,'0' as TD054,'0' as TD055,'' as TD056,'' as TD057
                                ,'' as TD058,'0' as TD059,'' as TD060,'0' as TD061,'' as TD062,'' as TD063,'' as TD064,'' as TD065
                                ,'' as TD066,'' as TD067,'' as TD068,'' as TD069
                                ,(select NN004 from CMSNN where NN001 = (select MA118 from COPMA where MA001 = [TMA001])) as TD070
                                ,'' as TD071,'' as TD072,'' as TD073,'' as TD074,'' as TD500,'0' as TD501,'' as TD502,'' as TD503
                                ,'' as TD504,'' as TD200,[ETA] as TD201,'0' as TD202,'' as TD203,'Y' as TD204,'' as TD205
                            FROM [CT_AUO_TOERP_Temp]
                                left join INVMB on INVMB.MB001 = [ERPNO]
                                left join INVMD on INVMD.MD001 = [ERPNO]
                                left join COPMG on MG001 = [TMA001] and MG002 = [ERPNO] and MG001 = [TMA001] 
		                            and COPMG.CREATE_DATE = (select max(CREATE_DATE) from COPMG where MG001 = [TMA001] and MG002 = [ERPNO] and MG001 = [TMA001] ) 
                                left join COPMA on MA001 = [TMA001]
                            where [TTC002] ='{0}'
                                order by TD002,TD003;", str_ERPNum_TC002);

            this.sqlDataAdapter1.SelectCommand.CommandText = str_sql_td;
            this.sqlDataAdapter1.Fill(dt_單身);
            //MyCode.Sql_dt(str_sql_td, dt_單身);

            this.get_total(dt_單身, str_ERPUP_CustID, str_ERPNum_TC002);

            //int_CheckError_TOERP_Temp = 0;
        }

        //TODO:計算 單頭 總金額.總數量.總包裝數
        private void get_total(DataTable dt,string str_total_TMA001,string str_total_ERPNum_TC002)
        {
            string str_sql_column_c = "", str_sql_value_c = "", str_sql_columns_c = "", str_sql_values_c = "";
            string str_sql_column_d = "", str_sql_value_d = "", str_sql_columns_d = "", str_sql_values_d = "";
            string data_type_d = "", data_type_c = "";
            bool bol_to_insert = false;
            this.ft_sum採購金額 = 0; this.ft_sum數量合計 = 0; this.ft_sum包裝數量合計 = 0;

            //// 準備ERP上傳 字串
            //str_sql_coptc = "";
            //str_sql_coptd = "";

            //TODO: 取得 COPTC.COPTD 客戶訂單單頭單身資料檔的欄位資料型態
            DataTable dt_schema_c = new DataTable();
            DataTable dt_schema_d = new DataTable();

            string str_sqlschema_c = String.Format(@"
                select COLUMN_NAME,DATA_TYPE,IS_NULLABLE
                from INFORMATION_SCHEMA.COLUMNS
               where TABLE_NAME='COPTC'");
            this.sqlDataAdapter1.SelectCommand.CommandText = str_sqlschema_c;
            this.sqlDataAdapter1.Fill(dt_schema_c);
            //MyCode.Sql_dt(str_sqlschema_c, dt_schema_c);

            string str_sqlschema_d = String.Format(@"
                select COLUMN_NAME,DATA_TYPE,IS_NULLABLE
               from INFORMATION_SCHEMA.COLUMNS
                where TABLE_NAME='COPTD'");
            this.sqlDataAdapter1.SelectCommand.CommandText = str_sqlschema_d;
            this.sqlDataAdapter1.Fill(dt_schema_d);
            //MyCode.Sql_dt(str_sqlschema_d, dt_schema_d);

            //TODO: 填入[COMPANY],[CREATOR],[USR_GROUP] ,[CREATE_DATE] ,[MODIFIER],[MODI_DATE] ,[FLAG]
            string[] str_basic =
                {
                    this.str_廠別,
                    this.str_ERPUP建立者ID,
                    this.str_ERPUP建立者GP,
                    this.str_ERPUP建立日期,
                    "",
                    "",
                    "1"
                };
            //TODO: 交易機制-單頭及單身寫入
            //20210816 須注意 有使用[SAP].[dbo].fm_COPTC_log、[SAP].[dbo].fm_COPTD_log，同ERP欄位 前面加入 DEL_DATE
            using (TransactionScope scope1 = new TransactionScope())
            {
                try
                {
                    sql_del_COPTCD = String.Format(
@"INSERT [SAP].[dbo].fm_COPTC_log
select GETDATE(),* from COPTC
where TC001 = '223' and TC002 = '{0}' and TC032 = '{1}' 

INSERT [SAP].[dbo].fm_COPTD_log
select GETDATE(),* from COPTD
where TD001 = '223' and TD002 = '{0}'

delete COPTC
where TC001 = '223' and TC002 = '{0}' and TC032 = '{1}' 

delete COPTD
where TD001 = '223' and TD002 = '{0}'", str_total_ERPNum_TC002, str_total_TMA001);

                    sql_del_tran_COPTCD += sql_del_COPTCD + "\r\n";

                    //MyCode.sqlExecuteNonQuery(sql_del_COPTCD);
                    this.to_ExecuteNonQuery(sql_del_COPTCD);

                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        for (j = 0; j < dt_schema_d.Rows.Count; j++)
                        {
                            bol_to_insert = false;
                            str_sql_column_d = dt_schema_d.Rows[j]["COLUMN_NAME"].ToString().Trim();
                            data_type_d = dt_schema_d.Rows[j]["DATA_TYPE"].ToString().Trim();
                            switch (str_sql_column_d.Substring(0, 3))
                            {
                                case "TD0":
                                case "TD5":
                                case "TD2":
                                    if (dt.Columns.Contains(str_sql_column_d) == true)
                                    {
                                        str_sql_value_d = this.get_sql_value(data_type_d, dt.Rows[i][str_sql_column_d].ToString().Trim());
                                        bol_to_insert = true;
                                    }
                                    break;

                                case "UDF":
                                    break;

                                default:
                                    //TODO: 填入[COMPANY],[CREATOR],[USR_GROUP] ,[CREATE_DATE] ,[MODIFIER],[MODI_DATE] ,[FLAG]
                                    if (j >= 0 && j <= 6)
                                    {
                                        //get_sql_value() 若為文字，前面補上N'，強制為string為Unicode字符串
                                        str_sql_value_d = this.get_sql_value(data_type_d, str_basic[j]);
                                        bol_to_insert = true;
                                    }

                                    break;
                            }
                            //TODO: 產生SQL字串
                            if (bol_to_insert == true)
                            {
                                str_sql_columns_d = str_sql_columns_d + str_sql_column_d + ",";
                                str_sql_values_d = str_sql_values_d + str_sql_value_d + ",";
                            }
                        }//for j

                        //新增至 COPTD	客戶訂單單身資料檔
                        //刪除最後的","符號
                        str_sql_columns_d = str_sql_columns_d.TrimEnd(new char[] { ',' });
                        str_sql_values_d = str_sql_values_d.TrimEnd(new char[] { ',' });
                        str_sql_d = String.Format(@"insert into COPTD ({0})
VALUES({1})", str_sql_columns_d, str_sql_values_d);

                        // 上傳ERP單身字串整理後，屆時一次上傳
                        str_sql_coptd += str_sql_d + "\r\n";
                        str_sql_tran_COPTC += str_sql_d + "\r\n";

                        this.to_ExecuteNonQuery(str_sql_d);
                        //MyCode.sqlExecuteNonQuery(str_sql_d);

                        //加總
                        ft_sum採購金額 += Convert.ToDouble(dt.Rows[i]["TD012"].ToString());
                        ft_sum數量合計 += Convert.ToDouble(dt.Rows[i]["TD008"].ToString()); ;
                        ft_sum包裝數量合計 += Convert.ToDouble(dt.Rows[i]["TD032"].ToString()); ;

                        //TODO:判別 單身單號與下一筆不同 則新增 單頭
                        //20210111 CONVERT(varchar(6) 改為 CONVERT(varchar(7)
                        // ", REPLICATE('0', (7 - LEN(CONVERT(varchar(6), (CFIPO.ERP_Num + " + str_key_客訂單號 + "))))) +CONVERT(varchar(6), (CFIPO.ERP_Num + " + str_key_客訂單號 + ")) as TC002" + str_enter +
                        if ((i != dt.Rows.Count - 1 && (dt.Rows[i]["TD002"].ToString() != dt.Rows[i + 1]["TD002"].ToString())) || i == dt.Rows.Count - 1)
                        {
                            DataTable dt_單頭 = new DataTable();
                            //TC027 確認碼
                            string str_sql_tc = String.Format(@"select * from (select [TTC001] as TC001,[TTC002] as TC002
                                ,[TTC039] as TC003 ,[TMA001] as TC004,'' as TC005
                                ,COPMA.MA016 as TC006,'002' as TC007,COPMA.MA014 as TC008 
                                ,(select MG004 from CMSMG where MG001 = COPMA.MA014 and MG002 = (select MAX(MG002) from CMSMG where MG001 = COPMA.MA014))  as TC009
                                ,(COPMA.MA080 + ' ' +COPMA.MA027) as TC010,COPMA.MA064 as TC011,'' as TC012,COPMA.MA030 as TC013,COPMA.MA031 as TC014
                                ,'' as TC015,COPMA.MA038 as TC016,'' as TC017,COPMA.MA005 as TC018,COPMA.MA048 as TC019,'' as TC020
                                ,'' as TC021,COPMA.MA056 as TC022,COPMA.MA057 as TC023,COPMA.MA058 as TC024,'' as TC025,COPMA.MA059 as TC026
                                ,'Y' as TC027,'0' as TC028,'0' as TC029,'0' as TC030,sum_AMOUNT as TC031
                                ,[TMA001] as TC032,'' as TC033,'' as TC034,COPMA.MA051 as TC035,'' as TC036,'' as TC037,'' as TC038
                                ,[TTC039] as TC039,'' as TC040
                                ,(select NN004 from CMSNN where NN001 = (select MA118 from COPMA where MA001 = [TMA001]))  as TC041
                                ,COPMA.MA083 as TC042,'0' as TC043,'0' as TC044,COPMA.MA095 as TC045,'0' as TC046
                                ,'' as TC047,'N' as TC048,'' as TC049,'N' as TC050,'' as TC051,'0' as TC052,COPMA.MA003 as TC053
                                ,'' as TC054,'' as TC055,'1' as TC056,'N' as TC057,'' as TC058,'' as TC059,'N' as TC060,'' as TC061
                                ,'' as TC062,(COPMA.MA079 + ' ' + COPMA.MA025) as TC063,COPMA.MA026 as TC064,COPMA.MA003 as TC065,COPMA.MA006 as TC066
                                ,COPMA.MA008 as TC067,'1' as TC068,'0000' as TC069,'N' as TC070,COPMA.MA110 as TC071,'0' as TC072
                                ,'0' as TC073,'' as TC074,'' as TC075,'' as TC076,'N' as TC077,COPMA.MA118 as TC078,'' as TC079
                                ,'' as TC080,'' as TC081,COPMA.MA076 as TC082,COPMA.MA077 as TC083,COPMA.MA078 as TC084,'' as TC085
                                ,'' as TC086,'' as TC087,'' as TC088,'' as TC089,'' as TC090,COPMA.MA123 as TC091,'' as TC092,'' as TC200
                                from (
                                    SELECT [TTC001],[TTC002],[Ver],[TTC039],sum([AMOUNT]) as 'sum_AMOUNT',[TMA001],[TCOMPANY],[TCREATOR],[TUSR_GROUP],[TCREATE_DATE] 
                                    FROM [CT_AUO_TOERP_Temp]  
                                    group by [TTC001],[TTC002],[Ver],[TTC039],[TMA001],[TCOMPANY],[TCREATOR],[TUSR_GROUP],[TCREATE_DATE])[CT_AUO_TOERP_Temp]
                                    left join COPMA on '{0}' = COPMA.MA001) CT_AUO_TOERP_Temp where [TC002] ='{1}'", str_total_TMA001, str_total_ERPNum_TC002);

                            this.sqlDataAdapter1.SelectCommand.CommandText = str_sql_tc;
                            this.sqlDataAdapter1.Fill(dt_單頭);
                            //MyCode.Sql_dt(str_sql_tc, dt_單頭);

                            for (x = 0; x < dt_單頭.Rows.Count; x++)
                            {
                                for (y = 0; y < dt_schema_c.Rows.Count; y++)
                                {
                                    bol_to_insert = false;
                                    str_sql_column_c = dt_schema_c.Rows[y]["COLUMN_NAME"].ToString().Trim();
                                    data_type_c = dt_schema_c.Rows[y]["DATA_TYPE"].ToString().Trim();

                                    switch (str_sql_column_c.Substring(0, 3))
                                    {
                                        case "TC0":
                                        case "TC5":
                                        case "TC2":
                                            if (dt_單頭.Columns.Contains(str_sql_column_c) == true)
                                            {
                                                str_sql_value_c = this.get_sql_value(data_type_c, dt_單頭.Rows[x][str_sql_column_c].ToString().Trim());

                                                bol_to_insert = true;

                                            }
                                            break;

                                        case "UDF":
                                            break;

                                        default:
                                            //TODO: 填入[COMPANY],[CREATOR],[USR_GROUP] ,[CREATE_DATE] ,[MODIFIER],[MODI_DATE] ,[FLAG]
                                            if (y >= 0 && y <= 6)
                                            {
                                                //get_sql_value() 若為文字，前面補上N'，強制為string為Unicode字符串
                                                str_sql_value_c = this.get_sql_value(data_type_c, str_basic[y]);
                                                bol_to_insert = true;
                                            }

                                            break;
                                    }
                                    //TODO: 產生SQL字串
                                    if (bol_to_insert == true)
                                    {
                                        str_sql_columns_c = str_sql_columns_c + str_sql_column_c + ",";
                                        str_sql_values_c = str_sql_values_c + str_sql_value_c + ",";
                                    }

                                }//for y

                                //新增至 COPTC 客戶訂單單頭資料檔
                                //刪除最後的","符號
                                str_sql_columns_c = str_sql_columns_c.TrimEnd(new char[] { ',' });
                                str_sql_values_c = str_sql_values_c.TrimEnd(new char[] { ',' });
                                str_sql_c =
                                    "insert into COPTC(" + str_sql_columns_c + ")" + str_enter +
                                    "VALUES(" + str_sql_values_c + ")";

                                // 上傳ERP單頭字串整理後，屆時一次上傳
                                str_sql_coptc += str_sql_c + str_enter;

                                this.to_ExecuteNonQuery(str_sql_c);
                                //MyCode.sqlExecuteNonQuery(str_sql_c);

                                //sqlapp log
                                str_sql_log = String.Format(
                                          @"insert into develop_app_log VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')"
                                          , str_ERPUP建立者ID, str_ERPUP建立日期, dt_單頭.Rows[x]["TC001"], dt_單頭.Rows[x]["TC002"], "COPTC", "fm_AUOPlannedOrder", "新增客戶計劃訂單單頭", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

                                // 上傳ERP Log字串整理後，屆時一次上傳
                                str_sql_logs += str_sql_log + str_enter;
                                str_sql_ListC += "'" + dt_單頭.Rows[x]["TC002"].ToString() + "',";

                                //已新增資料 重新累計
                                this.ft_sum包裝數量合計 = 0;
                                this.ft_sum採購金額 = 0;
                                this.ft_sum數量合計 = 0;

                                str_sql_columns_c = "";
                                str_sql_values_c = "";
                                str_sql_column_c = "";
                                str_sql_value_c = "";
                            } // for x

                        } // 單頭

                        str_sql_columns_d = "";
                        str_sql_values_d = "";
                        str_sql_column_d = "";
                        str_sql_value_d = "";

                    }//for i

                    //scope.Complete();

                    //列出轉換成ERP格式 單頭及單身
                    DataTable dt_coptd = new DataTable();
                    this.sqlDataAdapter1.SelectCommand.CommandText =
                        "select TD001 as 單別 ,TD002 as 單號 ,TD003 as 序號 ,TD004 as 品號 ,TD005 as 品名,TD006 as 規格 " +
                        ",TD007 as 庫別 ,TD010 as 單位 ,TD011 as 單價 ,TD008 as 訂單數量,TD012 as 金額 ,TD036 as 包裝單位 " +
                        ",TD032 as 訂單包裝數量 ,TD014 as 客戶品號,TD020 as 備註 ,TD202 as 實際可交貨數,TD013 as 預交日 " +
                        ",TD047 as 原預交日 ,TD048 as 排定交貨日 ,TD201 as 希望交貨日 from COPTD " +
                        "where TD001 ='223' and TD002 in('" + dt.Rows[0]["TD002"].ToString() + "') order by TD002";
  
                    this.sqlDataAdapter1.Fill(dt_coptd);
                    dgv_COPTD.DataSource = dt_coptd;

                    dt_tran_COPTD.Merge(dt_coptd);

                    DataTable dt_coptc = new DataTable();
                    this.sqlDataAdapter1.SelectCommand.CommandText =
                        "select TC001 as 單別,TC002 as 單號,TC003 as 訂單日期,TC004 as 客戶代號,TC005 as 部門代號" +
                        ",TC006 as 業務人員,TC008 as 交易幣別,TC009 as 匯率,TC041 as 營業稅率,TC012 as 客戶單號" +
                        ",TC029 as 訂單金額,TC030 as 訂單稅額,TC031 as 總數量,TC046 as 總包裝數量,TC053 as 客戶全名" +
                        ",TC010 as 送貨地址_一,TC014 as 付款條件,TC018 as 連絡人 from COPTC " +
                        "where TC001 ='223' and TC002 in('" + dt.Rows[0]["TD002"].ToString() + "')  order by TC002";
                    
                    this.sqlDataAdapter1.Fill(dt_coptc);
                    dgv_COPTC.DataSource = dt_coptc;
                    //MyCode.Sql_dgv(str_sql, dt_coptc, dgv_tc);

                    dt_tran_COPTC.Merge(dt_coptc);

                    //lab_status.Text = " ERP格式轉換完成";
                    //tctl_ERPUP.SelectedIndex = 3;
                    //tctl_ERPUP.SelectedIndex = 2;

                    ////確認資料轉換成 ERP格式後，開啟 上傳ERP按鈕
                    //btn_erpup.Enabled = true;
                    //btn_erpup.BackColor = System.Drawing.Color.SteelBlue;
                    //btn_erpup.ForeColor = System.Drawing.Color.White;

                    //txterr.Text += Environment.NewLine +
                    //                DateTime.Now.ToString() + Environment.NewLine +
                    //               ">> ERP格式轉換完成" + Environment.NewLine +
                    //               "===========";

                    //scope.Complete();
                }
                catch (Exception ex)
                {
                    lab_status.Text = " 錯誤：請檢查檔案重新上傳!!";
                    MessageBox.Show("第" + (i + 1) + "筆 " + "，轉換ERP格式 失敗!!" + Environment.NewLine +
                                    "【 " + ex.Message + " 】" + Environment.NewLine +
                                    "請先檢查【來源Excel格式】重新上傳 或 連絡MIS", "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txterr_ERPUP.Text += Environment.NewLine +
                                    DateTime.Now.ToString() + Environment.NewLine +
                                   "來源 第" + (i + 1) + "筆 " + "，轉換ERP格式 失敗!!" + Environment.NewLine +
                                    "【 " + ex.Message + " 】" + Environment.NewLine +
                                   "請先檢查【來源Excel格式】重新上傳 或 連絡MIS" + Environment.NewLine +
                                   "===========";
                    
                    int_CheckError_TOERP_Temp += 1;

                    btn_toerp.Enabled = false;
                    btn_toerp.BackColor = System.Drawing.SystemColors.Control;
                    btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;

                    //關閉 上傳ERP按鈕
                    btn_erpup.Enabled = false;
                    btn_erpup.BackColor = System.Drawing.SystemColors.Control;
                    btn_erpup.ForeColor = System.Drawing.SystemColors.ControlText;
                    
                    //dgv_excel.CurrentCell = dgv_excel.Rows[i].Cells[1];

                    return;
                }
                //發生例外時，會自動rollback
                finally
                {
                    this.sqlConnection1.Close();
                }
            }
        }
        
        private void btn_erpup_Click(object sender, EventArgs e)
        {
            //TODO:交易機制-確認 資料無誤，上傳ERP系統
            using (TransactionScope scope = new TransactionScope())
            {
                try
                {
                    DialogResult Result = MessageBox.Show("請再次確認資料-"+ "\r\n" + str_sql_ListC.TrimEnd(',') , "確認上傳ERP", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                    if (Result == DialogResult.OK)
                    {
                        this.to_ExecuteNonQuery(sql_del_tran_COPTCD);
                        this.to_ExecuteNonQuery(str_sql_coptc);
                        this.to_ExecuteNonQuery(str_sql_coptd);
                        this.to_ExecuteNonQuery(str_sql_logs);

                        scope.Complete();

                        MessageBox.Show("已上傳至ERP系統，將清除[整理結果]", "清除[整理結果]", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2);

                        txterr_ERPUP.Text += Environment.NewLine +
                                    DateTime.Now.ToString() + Environment.NewLine +
                                   ">> 已上傳至ERP系統-" + str_sql_ListC.TrimEnd(',') + Environment.NewLine +
                                   "===========";
                        //TODO:上傳 ERP系統完成後，將單據號碼.單據日期.檔案路徑.EXCEL匯入.CFIPO畫面清除，
                        //並關閉 轉換ERP格式及上傳ERP按鈕
                        
                        int_check_ERPUP_OK = 0;

                        if (int_check_ERPUP_OK == 0)
                        {
                            Clean_ERPUP();
                            tctl_ERPUP.SelectedIndex = 0;
                            bool_changed_dgv_ERPUP_Edit = false;
                        }
                        
                    }
                    else if (Result == DialogResult.Cancel)
                    {
                        return;
                    }
                }
                catch (Exception ex)
                {
                    lab_status.Text = " 錯誤：請檢查 單頭.單身 檔案重新上傳!!";
                    txt_path.Text = "";
                    MessageBox.Show("【 " + ex.Message + " 】" + Environment.NewLine +
                                    "請先重新執行操作，重新上傳 或 連絡MIS", "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                    txterr.Text += Environment.NewLine +
                                   DateTime.Now.ToString() + Environment.NewLine +
                                   "【 " + ex.Message + " 】" + Environment.NewLine +
                                   "請先檢查【整理結果】重新上傳 或 連絡MIS" + Environment.NewLine +
                                   "===========";
                    return;
                }
                //發生例外時，會自動rollback
                finally
                {
                    this.sqlConnection1.Close();
                }
            }
        }

        //private void Check_ERPNum_Status(string str_ERPNum_Status, string str_ERPNum_TC002,ComboBox cbo_ERP_M)
        //{
        //    DataTable dt_OrderVer = new DataTable();
        //    string sql_OrderVer = String.Format(@"select TC069 as 版號 from COPTC where TC001='223' and TC002 = '{0}'", str_ERPNum_TC002);
        //    MyCode.Sql_dt(sql_OrderVer, dt_OrderVer);

        //    switch (str_ERPNum_Status)
        //    {
        //        case "新增":
        //            if (dt_OrderVer.Rows.Count > 0)
        //            {
        //                string str_ErrorMessage = "單號已使用-" + str_ERPNum_TC002 + "，無法新增，將改為【變更】狀態";
        //                MessageBox.Show(str_ErrorMessage);
        //                //MyCode.Error_MessageBar(txterr, str_ErrorMessage);
        //                //MessageBox.Show("單號已使用-" + str_ERPNum_TC002 + "，無法新增");
        //                cbo_ERP_M.Text = "變更";
        //            }
        //            break;
        //        case "變更":
        //            if (dt_OrderVer.Rows.Count == 0)
        //            {
        //                string str_ErrorMessage = "無符合單號-" + str_ERPNum_TC002 + "，無法變更，將改為【新增】狀態";
        //                int_CheckError_TOERP_Temp = 1;
        //                MessageBox.Show(str_ErrorMessage);
        //                //MyCode.Error_MessageBar(txterr, str_ErrorMessage);
        //                //MessageBox.Show("無符合單號-" + str_ERPNum_TC002 + "，無法變更");
        //                cbo_ERP_M.Text = "新增";
        //            }
        //            break;
        //        case "無":
        //            break;
        //    }
        //}

        private void cbo_ERPUP建立者_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.str_ERPUP建立者ID = this.dt_建立者.Rows[this.cbo_ERPUP建立者.SelectedIndex]["MF001"].ToString().Trim();
            this.str_ERPUP建立者GP = this.dt_建立者.Rows[this.cbo_ERPUP建立者.SelectedIndex]["MF004"].ToString().Trim();
        }

        private void cbo_ERPUP_Cust_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.str_ERPUP_CustID = this.dt_Cust.Rows[this.cbo_ERPUP_Cust.SelectedIndex]["MA001"].ToString().Trim();
        }

        private void button_單據日期_Click(object sender, EventArgs e)
        {
            //TODO:單頭及單身若不為空值，表示已轉換Excel格式，需重新轉換 或 資料已上傳資料庫，需重新選擇日期
            //資料上傳資料庫後，dgv_ImportExcel會清空
            if (dgv_ERPUP_Edit.Rows.Count != 0)
            {
                DialogResult Result = MessageBox.Show("修改 單據日期 後，需重新【查詢】", "已查詢", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                if (Result == DialogResult.OK)
                {
                    lab_ERPUP_status.Text = "請 選擇【單據日期】";
                    Clean_TOERP_Temp();

                    //MyCode.Error_MessageBar(txterr,"修改來源日期，請重新【選擇Excel檔案】");
                    //txterr.Text += Environment.NewLine +
                    //           DateTime.Now.ToString() + Environment.NewLine +
                    //           " 修改來源日期，請重新【選擇Excel檔案】" + Environment.NewLine +
                    //           "===========";

                    this.fm_月曆 = new 月曆(this.textBox_單據日期, this.button_單據日期, "單據日期");
                }
                else if (Result == DialogResult.Cancel)
                {
                    return;
                }
            }
            else
            {
                this.fm_月曆 = new 月曆(this.textBox_單據日期, this.button_單據日期, "單據日期");

            }
        }

        private void cbo_ERPUP_ShowM_SelectedValueChanged(object sender, EventArgs e)
        {
            //dgv_ERPUP_Edit.DataSource = null;
            //dt_AddERP_Search.Reset();
            Clean_TOERP_Temp();
        }

        //private void Error_MessageBar(string str_MessageBar)
        //{
        //    txterr.Text += Environment.NewLine +
        //                    DateTime.Now.ToString() + Environment.NewLine +
        //                    ">> " + str_MessageBar + Environment.NewLine +
        //                    "===========";
        //}

        private void Clean_TOERP_Temp()
        {
            cbo_ERPUP_Line.SelectedIndex = 0;
            
            dgv_ERPUP_Edit.DataSource = null;
            dgv_TOERP_Temp.DataSource = null;
            dgv_COPTC.DataSource = null;
            dgv_COPTD.DataSource = null;


            //dt_AddERP_Search.Reset();
            //dt_TOERP_Temp.Reset();
            dt_COPTC.Reset();
            dt_COPTD.Reset();

            btn_toerp.Enabled = false;
            btn_toerp.BackColor = System.Drawing.SystemColors.Control;
            btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;

            btn_erpup.Enabled = false;
            btn_erpup.BackColor = System.Drawing.SystemColors.Control;
            btn_erpup.ForeColor = System.Drawing.SystemColors.ControlText;
        }

        private void Clean_ERPUP()
        {
            int_check_ERPUP_OK = 1;
            //cbo_ERPUP_Line.SelectedIndex = 0;

            dgv_TOERP_Temp.DataSource = null;
            dgv_COPTC.DataSource = null;
            dgv_COPTD.DataSource = null;

            //dt_AddERP_Search.Reset();
            dt_COPTC.Reset();
            dt_COPTD.Reset();

            //btn_toerp.Enabled = false;
            //btn_toerp.BackColor = System.Drawing.SystemColors.Control;
            //btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;

            btn_erpup.Enabled = false;
            btn_erpup.BackColor = System.Drawing.SystemColors.Control;
            btn_erpup.ForeColor = System.Drawing.SystemColors.ControlText;
        }

        private void btn_fileopen_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process prc = new System.Diagnostics.Process();
            prc.StartInfo.FileName = txt_path.Text.ToString();
            prc.Start();
        }

        private void Custom_DTInputExcel(ClosedXML.Excel.IXLWorksheet wsheet, DataTable dt,string Line_Name)
        {
            //前回差20200814
            //0818FCST

            DataTable dt_ToExcel = new DataTable();
            dt_ToExcel = dt;
            //dt_ToExcel.Columns.RemoveAt(0);
            //dt_ToExcel.AcceptChanges();

            //因兩個線別合併，須拆分，以便查詢
            if (Line_Name == "CSDC6C") 
            {
                Line_Name = "C5D','C6C";
            }
            //判別專用料加入底色
            string sql_SPECIAL = String.Format(@"SELECT [ERP_NO],[SPECIAL] FROM [A01A].[dbo].[CT_AUO_ERPNO]
                             where  [SPECIAL] = 'Y' and [FAB] in ('{0}')", Line_Name);
            DataTable dt_SPECIAL = new DataTable();
            MyCode.Sql_dt(sql_SPECIAL, dt_SPECIAL);

            Dictionary<string, string> dict_SPECIAL = dt_SPECIAL.AsEnumerable()
                .ToDictionary<DataRow, string, string> (
                row => row.Field<string>("ERP_NO"),
                row => row.Field<string>("SPECIAL"));

            //設定數字顯示格式
            for (int k = 2; k < 10; k++)
            {
                wsheet.Column(k).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
            }

            //插入標題，並設定為文字型態
            int x = 0;
            wsheet.Row(2).Style.NumberFormat.Format = "@";
            //wsheet.Row(1).Style.Fill.BackgroundColor = XLColor.FromHtml("#3366FF");
            wsheet.Row(1).Style.Font.SetBold();
            wsheet.Row(2).Style.Font.SetBold();
            //wsheet.Row(1).Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            wsheet.Row(2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            wsheet.Row(1).Style.Font.FontColor = XLColor.FromHtml("#FFFFFF");
            wsheet.Row(2).Style.Font.FontColor = XLColor.FromHtml("#FFFFFF");


            wsheet.Column(3).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFFF99");
            wsheet.Column(5).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFFF99");
            wsheet.Column(7).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFFF99");
            wsheet.Column(9).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFFF99");


            //標題名稱
            //FFFF99 黃色 #3366FF 藍色
            foreach (DataColumn Column in dt_ToExcel.Columns)
            {
                wsheet.Cell(2, x + 1).Value = Column.ColumnName.ToString();
                wsheet.Cell(2, x + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#3366FF");
                wsheet.Cell(1, x + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#3366FF");
                if (Column.ColumnName.ToString().Substring(0,2) == "前回") 
                {
                    wsheet.Cell(2, x + 1).Value = "前回差";
                    wsheet.Cell(2, x + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#808080");
                }

                x++;
            }

            //標題文字
            wsheet.Cell(1, 1).Value = "前回差" + cbo_LastDate.Text.ToString();
            switch (Line_Name)
            {
                case "C4A":
                    wsheet.Cell(2, 1).Value = cbo_OrderLogDate.Text.ToString().Substring(4, 4) + "H10FCST";
                    break;
                case "CSDC6C":
                    wsheet.Cell(2, 1).Value = cbo_OrderLogDate.Text.ToString().Substring(4, 4) + "H11FCST";
                    break;
                case "C5E":
                    wsheet.Cell(2, 1).Value = cbo_OrderLogDate.Text.ToString().Substring(4, 4) + "H14FCST";
                    break;
            }

            //插入資料
            int i = 0;
            foreach (DataRow row in dt_ToExcel.Rows)
            {
                int j = 0;
                foreach (DataColumn Column in dt_ToExcel.Columns)
                {
                    wsheet.Cell(i + 3, j + 1).Value = row[j];
                    //如果是專用料，加註淺藍色底色
                    if (dict_SPECIAL.ContainsKey(row[j].ToString()) == true ) 
                    {
                        wsheet.Cell(i + 3, j + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#00ccff");
                    }
                    j++;
                }
                i++;

                if ( i >= dt_ToExcel.Rows.Count) 
                {
                    wsheet.Cell(i + 3, 3).FormulaA1 = "sum(C3:C" + (i + 2) +")";
                    wsheet.Cell(i + 3, 4).FormulaA1 = "sum(D3:D" + (i + 2) + ")";
                    wsheet.Cell(i + 3, 5).FormulaA1 = "sum(E3:E" + (i + 2) + ")";
                    wsheet.Cell(i + 3, 6).FormulaA1 = "sum(F3:F" + (i + 2) + ")";
                    wsheet.Cell(i + 3, 7).FormulaA1 = "sum(G3:G" + (i + 2) + ")";
                    wsheet.Cell(i + 3, 8).FormulaA1 = "sum(H3:H" + (i + 2) + ")";
                    wsheet.Cell(i + 3, 9).FormulaA1 = "sum(I3:I" + (i + 2) + ")";
                    wsheet.Cell(i + 3, 10).FormulaA1 = "sum(J3:J" + (i + 2) + ")";
                }
            }

            //wsheet_dcMPT.Cell(j + 7, 11).FormulaA1 =

            //設定框線
            wsheet.Range("A2:J" + (dt_ToExcel.Rows.Count+ 2)).Style
                    .Border.SetTopBorder(XLBorderStyleValues.Thin)
                    .Border.SetBottomBorder(XLBorderStyleValues.Thin)
                    .Border.SetLeftBorder(XLBorderStyleValues.Thin)
                    .Border.SetRightBorder(XLBorderStyleValues.Thin);
            //凍結視窗
            wsheet.SheetView.Freeze(2, 1);
        }


        //bool IsToForm1 = false; //紀錄是否要回到Form1
        protected override void OnClosing(CancelEventArgs e) //在視窗關閉時觸發
        {
            //DialogResult dr = MessageBox.Show("\"是\"回到主畫面 \r\n \"否\"關閉程式", "是否要關閉程式"
            //    , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            //if (dr == DialogResult.Yes)
            //{
            //    IsToForm1 = true;
            //}

            //base.OnClosing(e);
            //if (IsToForm1) //判斷是否要回到Form1
            //{
            //    this.DialogResult = DialogResult.Yes; //利用DialogResult傳遞訊息
            //    fm_menu fm_menu = (fm_menu)this.Owner; //取得父視窗的參考
            //}
            //else
            //{
            //    this.DialogResult = DialogResult.No;
            //}
            if (bool_changed_dgv_ERPUP_Edit == true)
            {
                DialogResult Result = MessageBox.Show("已有編輯紀錄，是否要離開?", "警示", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                if (Result == DialogResult.Yes)
                {
                    Environment.Exit(Environment.ExitCode);
                }
                else
                {
                    e.Cancel = true; //取消關閉
                }

            }
            else 
            {
                Environment.Exit(Environment.ExitCode);
            }
        }


        public static void WriteLog(string message)
        {
            string DIRNAME = Application.StartupPath + @"\Log\";
            string FILENAME = DIRNAME + DateTime.Now.ToString("yyyyMMdd") + ".txt";

            if (!Directory.Exists(DIRNAME))
                Directory.CreateDirectory(DIRNAME);

            if (!File.Exists(FILENAME))
            {
                // The File.Create method creates the file and opens a FileStream on the file. You neeed to close it.
                File.Create(FILENAME).Close();
            }
            using (StreamWriter sw = File.AppendText(FILENAME))
            {
                Log(message, sw);
            }
        }

        private static void Log(string logMessage, TextWriter w)
        {
            w.Write("Log Entry : ");
            w.WriteLine("{0} {1}", DateTime.Now.ToLongDateString(), DateTime.Now.ToLongTimeString());
            w.WriteLine("手動備份:{0}", logMessage);
            w.WriteLine("-------------------------------");
        }

        
    }
}
