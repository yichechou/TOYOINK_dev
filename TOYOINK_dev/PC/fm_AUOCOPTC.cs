﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Transactions;
using Myclass;

namespace TOYOINK_dev
{
    /*20210111 CONVERT(varchar(6) 改為 CONVERT(varchar(7)
    *  ", REPLICATE('0', (7 - LEN(CONVERT(varchar(6), (CFIPO.ERP_Num + " + str_key_客訂單號 + "))))) +CONVERT(varchar(6), (CFIPO.ERP_Num + " + str_key_客訂單號 + ")) as TD002" + str_enter +
    * 20210401 ,COPMA.MA024 as TC063 欄位值錯誤，改為,COPMA.MA025 as TC063 發票地址(一)
    * 20210623 將建立者欄位改為鎖定，無法變更
    * 20210913 升級GP4單身增加兩個欄位，計價數量(TD076).計價單位(TD077)，同原欄位 數量(TD008).單位(TD010) 及更新連線方式改由MyClass代入
    *  ,CFIPO.Quantity as TD076,CFIPO.UOM as TD077,
    * 20230803 生管 林玲禎提出，EXCEL匯入時，檢查是否空值，忽略欄位名稱為[Sample]或[Remark]
    * 20240222 生管 林玲禎提出，1.EXCEL匯入時，僅檢查欄位名稱為線別,Number,Item,Item Description,UOM,Quantity,Currency,Need By Date 是否空值，
    *  2.檢查EXCEL金額[Shipment Amount]是否與ERP相符但不卡控[轉換ERP格式]僅[緊示]；3.新增[不轉換EXCEL客戶單號]選項
    * 20240229 生管 林玲禎提出，財務聯絡因財務報表製作時需要［部門別］資訊 COPMA.MA015對應訂單單頭COPTC.TC005
    * 20240305 生管 林玲禎提出 單價欄位比對錯誤(單價-TD011)，正確為金額[TD012]，(COPMB.MB008 * CFIPO.Quantity) as TD012
    * 
    */

    public partial class fm_AUOCOPTC : Form
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        /// 
        public MyClass MyCode;

        DataTable dt_COPMA = new DataTable();
        string str_sql = "", str_sql_c = "", str_sql_d = "", str_sql_coptc = "", str_sql_coptd = "";
        string str_sql_log = "", str_sql_logs = "";
        string str_enter = ((char)13).ToString() + ((char)10).ToString();
        string str_key_客訂單號 = "";
        string errtable = "";
        string str_廠別 = "A01A", str_建立者ID = "", str_建立者GP = "", str_建立日期 = "";
        月曆 fm_月曆;
        int i, j, x, y;
        DataTable dt_建立者;
        double ft_sum採購金額 = 0, ft_sum數量合計 = 0, ft_sum包裝數量合計 = 0;

        //TODO: 右上角訊息視窗，自動捲動置底
        private void txterr_TextChanged(object sender, EventArgs e)
        {
            txterr.SelectionStart = txterr.Text.Length;
            txterr.ScrollToCaret();  //跳到遊標處 
        }
        public fm_AUOCOPTC()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();

            //MyCode.strDbCon = MyCode.strDbConLeader;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConLeader;

            MyCode.strDbCon = MyCode.strDbConA01A;
            this.sqlConnection1.ConnectionString = MyCode.strDbConA01A;

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

        //TODO:建立者下拉式清單
        private void cob_建立者_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.str_建立者ID = this.dt_建立者.Rows[this.cob_建立者.SelectedIndex]["MF001"].ToString().Trim();
            this.str_建立者GP = this.dt_建立者.Rows[this.cob_建立者.SelectedIndex]["MF004"].ToString().Trim();
        }

        private void textBox_單據日期_TextChanged(object sender, EventArgs e)
        {
            //資料上傳ERP後，textBox_單據日期 會清空，需重新選擇
            if (string.IsNullOrEmpty(textBox_單據日期.Text))
            {
                return;
            }

            string num2_ym = "";
            string now_ym = "";
            string txt_date = textBox_單據日期.ToString();

            DateTime num2_date = DateTime.ParseExact((textBox_單據日期.Text.ToString()), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture);

            TaiwanCalendar nowdate = new TaiwanCalendar();
            //顯示民國日期格式為eee/mm/dd(105/09/29)  
            //月份或日期為1位數，想在前面補0湊成2位數，這種方法為PadLeft(2,'0')。
            now_ym = nowdate.GetYear(num2_date).ToString() + nowdate.GetMonth(num2_date).ToString().PadLeft(2, '0');
            num2_ym = now_ym.Substring(1, 4);

            //TODO:查詢最後一筆ERP使用單號
            str_sql =
                "select top 1 TC002 from COPTC" + str_enter +
                "where TC001 = '220'" + "and TC002 like '" + num2_ym.ToString() + "%'" + str_enter +
                "order by TC002 desc";
            this.sqlDataAdapter1.SelectCommand.CommandText = str_sql;
            DataTable dt_temp = new DataTable();
            this.sqlDataAdapter1.Fill(dt_temp);
            //MyCode.Sql_dt(str_sql, dt_temp);

            if (dt_temp.Rows.Count == 0)
            {
                str_key_客訂單號 = num2_ym.ToString() + "001";
                lab_num2.Text = this.str_key_客訂單號.PadLeft(7, '0');
            }
            else
            {
                this.str_key_客訂單號 = (Convert.ToInt32(dt_temp.Rows[0][0].ToString()) + 1).ToString();
                this.lab_num2.Text = this.str_key_客訂單號.PadLeft(7, '0');
            }
        }

        private void fm_AUOCOPTC_Load(object sender, EventArgs e)
        {
            //TODO:匯入ERP 可建立客戶訂單 使用者清單
            dt_建立者 = new DataTable();
            this.sqlDataAdapter1.SelectCommand.CommandText = "select MF001,MF001 + MF002 as 人員,MF002,MF004 from ADMMF";
            this.sqlDataAdapter1.Fill(dt_建立者);

            //str_sql = "select MF001,MF001 + MF002 as 人員,MF002,MF004 from ADMMF";
            //MyCode.Sql_dt(str_sql, dt_建立者);

            this.cob_建立者.Items.Clear();

            string str_建立者 = "";
            int check = 0;


            for (int i = 0; i < dt_建立者.Rows.Count; i++)
            {
                str_建立者 = this.dt_建立者.Rows[i]["MF002"].ToString().Trim();
                this.cob_建立者.Items.Add(dt_建立者.Rows[i]["人員"].ToString().Trim());

                if (str_建立者 == loginName || loginName == "周怡甄")
                {
                    this.cob_建立者.SelectedIndex = i;
                    check = 1;
                }

            }
            if (check == 0)
            {
                MessageBox.Show("非採購人員不能使用", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txterr.Text += Environment.NewLine +
                           DateTime.Now.ToString() + Environment.NewLine +
                           "非採購人員不能使用" + Environment.NewLine +
                           "===========";
                btn_file.Enabled = false;
                button_單據日期.Enabled = false;
                cob_建立者.Enabled = false;

                //fm_login fm_login = new fm_login();

                //fm_login.Show();
                //this.Hide();
                return;
            }

            //TODO:格式化 建立日期
            lab_Nowdate.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox_單據日期.Text = DateTime.Now.ToString("yyyyMMdd");
            str_建立日期 = lab_Nowdate.Text.ToString().Trim();
        }

        private void button_單據日期_Click(object sender, EventArgs e)
        {
            //TODO:單頭及單身若不為空值，表示已轉換ERP格式，需重新轉換 或 資料已上傳ERP，需重新選擇日期
            //資料上傳ERP後，dgv_excel會清空
            //if (dgv_tc.DataSource != null || dgv_td.DataSource != null || dgv_excel.DataSource != null)
            if (btn_toerp.Enabled == true || btn_erpup.Enabled == true )

            {
                DialogResult Result = MessageBox.Show("修改 單據日期 後，需重新【選擇檔案】", "Excel檔案已匯入", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                if (Result == DialogResult.OK)
                {
                    lab_status.Text = "請 選擇檔案";
                    txt_path.Text = "";
                    dgv_excel.DataSource = null;
                    dgv_cfipo.DataSource = null;
                    tabCtl_data.SelectedIndex = 0;
                    btn_toerp.Enabled = false;
                    btn_toerp.BackColor = System.Drawing.SystemColors.Control;
                    btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
                    btn_erpup.Enabled = false;
                    btn_erpup.BackColor = System.Drawing.SystemColors.Control;
                    btn_erpup.ForeColor = System.Drawing.SystemColors.ControlText;
                    dgv_tc.DataSource = null;
                    dgv_td.DataSource = null;

                    txterr.Text += Environment.NewLine +
                               DateTime.Now.ToString() + Environment.NewLine +
                               " 修改單據日期，請重新【選擇檔案】" + Environment.NewLine +
                               "===========";

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
                btn_file.Enabled = true;
                lab_status.Text = "請 選擇檔案";

            }

        }

       

        private void btn_file_Click(object sender, EventArgs e)
        {
            //TODO:判別 已轉換ERP格式，重新選擇檔案，需重新手動轉換ERP格式
            if ((dgv_tc.DataSource != null || dgv_td.DataSource != null) && dgv_excel.DataSource != null)
            {
                DialogResult Result = MessageBox.Show("需 重新執行【ERP格式轉換】", "已轉換 ERP格式", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                if (Result == DialogResult.OK)
                {
                    //TODO:關閉 ERP上傳 按鈕及清空 已轉換ERP格式的單頭及單身
                    btn_erpup.Enabled = false;
                    btn_erpup.BackColor = System.Drawing.SystemColors.Control;
                    btn_erpup.ForeColor = System.Drawing.SystemColors.ControlText;
                    dgv_tc.DataSource = null;
                    dgv_td.DataSource = null;
                }
                else if (Result == DialogResult.Cancel)
                {
                    return;
                }
            }
            else
            {
                lab_status.Text = "請 選擇檔案";
                dgv_tc.DataSource = null;
                dgv_td.DataSource = null;
            }

            //this.openFileDialog1.InitialDirectory = @"P:\共用區\生產關係\受発注管理\原物料發注資料\";

            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txt_path.Text = this.openFileDialog1.FileName;
            }
            else
            {
                return;
            }

            dgv_excel.DataSource = MyClass.ReadExcelToTable("fm_AUOCOPTC",txt_path.Text.ToString(),"1=1");

            DataTable dt_Excel訂單 = new DataTable();
            dt_Excel訂單 = (DataTable)this.dgv_excel.DataSource;

            // 20230803 檢查匯入的EXCEL是否有空值，如有空值，記錄位置
            List<Tuple<int, string>> emptyFields = new List<Tuple<int, string>>();

            //for (int i = 0; i < dt_訂單.Rows.Count; i++)
            //{
            //    DataRow row = dt_訂單.Rows[i];
            //    for (int j = 0; j < dt_訂單.Columns.Count; j++)
            //    {
            //        DataColumn col = dt_訂單.Columns[j];
            //        // 忽略欄位名稱為[Sample]或[Remark]
            //        if (col.ColumnName == "Sample" || col.ColumnName == "Remark")
            //            continue;


            //        // 檢查每個欄位的值是否為 DBNull.Value 或空值
            //        if (row.IsNull(col) || string.IsNullOrWhiteSpace(row[col.ColumnName].ToString()))
            //        {
            //            emptyFields.Add(new Tuple<int, string>(i, col.ColumnName));
            //        }
            //    }
            //}

            //20240222 生管 林玲禎提出 僅檢查以下欄位是否為空值
            var checkColumns = new HashSet<string>
            {
                "線別",
                "Number",
                "Item",
                "Item Description",
                "UOM",
                "Quantity",
                "Currency",
                "Need By Date"
            };

            for (int i = 0; i < dt_Excel訂單.Rows.Count; i++)
            {
                DataRow row = dt_Excel訂單.Rows[i];
                for (int j = 0; j < dt_Excel訂單.Columns.Count; j++)
                {
                    DataColumn col = dt_Excel訂單.Columns[j];
                    // 僅對指定欄位進行檢查
                    if (checkColumns.Contains(col.ColumnName))
                    {
                        // 檢查指定欄位的值是否為 DBNull.Value 或空值
                        if (row.IsNull(col) || string.IsNullOrWhiteSpace(row[col.ColumnName].ToString()))
                        {
                            emptyFields.Add(new Tuple<int, string>(i, col.ColumnName));
                        }
                    }
                    // 其他欄位不進行空值檢查，所以這裡不需要做任何操作
                }
            }


            if (emptyFields.Count > 0)
            {
                // 若有空值，顯示訊息和定位
                string errorMessage = "警告：請檢查以下欄位的空值重新上傳：" + Environment.NewLine;
                foreach (var field in emptyFields)
                {
                    errorMessage += $"【來源檔案-[{field.Item2}]】在第 {field.Item1 + 1} 列" + Environment.NewLine;
                }
                lab_status.Text = errorMessage;
                txt_path.Text = "";
                MessageBox.Show("轉換失敗!!" + Environment.NewLine +
                                errorMessage +
                                "請先檢查相應欄位重新上傳 或 連絡MIS", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                txterr.Text += Environment.NewLine +
                               DateTime.Now.ToString() + Environment.NewLine +
                               "轉換失敗!!" + Environment.NewLine +
                               errorMessage +
                               "===========";

                dgv_cfipo.DataSource = null;
                tabCtl_data.SelectedIndex = 0;
                btn_toerp.Enabled = false;
                btn_toerp.BackColor = System.Drawing.SystemColors.Control;
                btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
                dgv_excel.CurrentCell = dgv_excel.Rows[emptyFields[0].Item1].Cells[emptyFields[0].Item2];

                return;
            }

            //20240222 依客戶單號排序，再轉換
            // 假設 dt_訂單 是您的 DataTable 對象
            // 使用 DefaultView 來排序並建立新的 DataView
            DataView dv = dt_Excel訂單.DefaultView;
            // 根據 Number 欄位進行排序
            dv.Sort = "Number ASC"; // ASC 代表升序，DESC 代表降序
                                    // 將排序後的結果放回 DataTable
            DataTable dt_訂單 = dv.ToTable();

            this.lab_status.Text = " 轉換中，請稍後";
            this.to_ExecuteNonQuery("delete from CFIPO"); //MIS建立的暫存table
            //MyCode.sqlExecuteNonQuery("delete from CFIPO");

            string str_sql_column = "", str_sql_value = "", str_sql_columns = "", str_sql_values = "";
            string data_type = "";

            DataTable dt_schema = new DataTable();

            //TODO: 取得CFIPO單身資料檔的欄位資料型態
            string str_sql =
                "select COLUMN_NAME,DATA_TYPE,IS_NULLABLE" + str_enter +
                "from INFORMATION_SCHEMA.COLUMNS" + str_enter +
                "where TABLE_NAME='CFIPO'";
            this.sqlDataAdapter1.SelectCommand.CommandText = str_sql;
            this.sqlDataAdapter1.Fill(dt_schema);
            //MyCode.Sql_dt(str_sql, dt_schema);

            //TODO:交易機制-將Excel存入 CFIPO 資料表
            using (TransactionScope scope = new TransactionScope())
            {
                try
                {
                    int x = 1;
                    int y = 1;
                    for (i = 0; i < dt_訂單.Rows.Count; i++)
                    {
                        for (j = 0; j < dt_schema.Rows.Count; j++)
                        {
                            str_sql_column = dt_schema.Rows[j]["COLUMN_NAME"].ToString().Trim();
                            data_type = dt_schema.Rows[j]["DATA_TYPE"].ToString().Trim();

                            switch (j)
                            {
                                //ERP 號碼
                                case 0:
                                    if (i == 0)
                                    {
                                        str_sql_value = this.get_sql_value(data_type, x.ToString());
                                    }
                                    else if ((dt_訂單.Rows[i][2].ToString().Trim()) == (dt_訂單.Rows[i - 1][2].ToString().Trim()))
                                    {
                                        str_sql_value = this.get_sql_value(data_type, x.ToString());
                                    }
                                    else
                                    {
                                        x += 1;
                                        str_sql_value = this.get_sql_value(data_type, x.ToString());
                                    }
                                    break;

                                //序號
                                case 1:
                                    if (i == 0)
                                    {
                                        str_sql_value = this.get_sql_value(data_type, y.ToString().PadLeft(4, '0').ToString());
                                    }
                                    else if ((dt_訂單.Rows[i][2].ToString().Trim()) == (dt_訂單.Rows[i - 1][2].ToString().Trim()))
                                    {
                                        y += 1;
                                        str_sql_value = this.get_sql_value(data_type, (y.ToString()).PadLeft(4, '0'));
                                    }
                                    else
                                    {
                                        y = 1;
                                        str_sql_value = this.get_sql_value(data_type, (y.ToString()).PadLeft(4, '0'));
                                    }
                                    break;

                                //ERP 客戶代號
                                case 2:
                                    if ((dt_訂單.Rows[i][0].ToString().Trim()) == "C5E")
                                    {
                                        str_sql_value = this.get_sql_value(data_type, "AU-TK");
                                    }
                                    else
                                    {
                                        str_sql_value = this.get_sql_value(data_type, "AU-TN");
                                    }
                                    break;

                                //20240222 生管 林玲禎提出 不轉換EXCEL客戶單號判別
                                //ERP 客戶單號
                                case 3:
                                    if (chkNoTranNum.Checked == false)
                                    {
                                        if ((dt_訂單.Rows[i][0].ToString().Trim()) == "C5E")
                                        {
                                            //[線別]+'-'+[Number]+'-HC' 
                                            str_sql_value = this.get_sql_value(data_type, (dt_訂單.Rows[i][0].ToString().Trim() + '-' + dt_訂單.Rows[i][2].ToString().Trim() + "-HC"));
                                        }
                                        else
                                        {
                                            str_sql_value = this.get_sql_value(data_type, (dt_訂單.Rows[i][0].ToString().Trim() + '-' + dt_訂單.Rows[i][2].ToString().Trim() + "-LT"));
                                        }
                                    }
                                    else
                                    {
                                        str_sql_value = this.get_sql_value(data_type, (dt_訂單.Rows[i][2].ToString().Trim()));
                                    }

                                    //判別 客戶單號有沒有重複
                                    string str_TC012 = str_sql_value.Substring(1);

                                    DataTable dt_TC012 = new DataTable();
                                    string str_sql_TC012 = "select TC001,TC002,TC004,TC012 from COPTC where TC012 = " + str_TC012;

                                    this.sqlDataAdapter1.SelectCommand.CommandText = str_sql_TC012;
                                    this.sqlDataAdapter1.Fill(dt_TC012);
                                    //MyCode.Sql_dt(str_sql_TC012, dt_TC012);

                                    if (dt_TC012.Rows.Count != 0)
                                    {
                                        lab_status.Text = " 警告：請檢查【來源檔案-[Number]】重新上傳!!";
                                        txt_path.Text = "";
                                        MessageBox.Show("來源 第" + (i + 1) + "筆 " + "，轉換失敗!!" + Environment.NewLine +
                                                        "與【單據：" + dt_TC012.Rows[0]["TC001"].ToString().Trim() + "-" + dt_TC012.Rows[0]["TC002"].ToString().Trim() + "】，" + Environment.NewLine +
                                                        "【客戶單號：" + str_TC012 + "】重複!!" + Environment.NewLine +
                                                        "請先檢查【來源Excel-[Number]欄位】重新上傳 或 連絡MIS", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                                        txterr.Text += Environment.NewLine +
                                                       DateTime.Now.ToString() + Environment.NewLine +
                                                       "來源 第" + (i + 1) + "筆 " + "，轉換失敗!!" + Environment.NewLine +
                                                        "與【單據：" + dt_TC012.Rows[0]["TC001"].ToString().Trim() + "-" + dt_TC012.Rows[0]["TC002"].ToString().Trim() + "】，" +
                                                       "【客戶單號：" + str_TC012 + "】重複!!" + Environment.NewLine +
                                                       "請先檢查【來源Excel-[Number]欄位】重新上傳 或 連絡MIS" + Environment.NewLine +
                                                       "===========";

                                        dgv_cfipo.DataSource = null;
                                        tabCtl_data.SelectedIndex = 0;
                                        btn_toerp.Enabled = false;
                                        btn_toerp.BackColor = System.Drawing.SystemColors.Control;
                                        btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
                                        dgv_excel.CurrentCell = dgv_excel.Rows[i].Cells["Number"];

                                        return;
                                    }
                                    break;

                                //ERP 幣別
                                case 4:
                                    if (dt_訂單.Rows[i]["Currency"].ToString().Trim() == "TWD")
                                    {
                                        str_sql_value = this.get_sql_value(data_type, "NTD");
                                    }
                                    else
                                    {
                                        lab_status.Text = " 警告：請檢查【來源檔案-[Currency]】重新上傳!!";
                                        MessageBox.Show("來源 第" + (i + 1) + "筆 " + "，轉換失敗!!" + Environment.NewLine +
                                                        "幣別 不等於 TWD 無法轉換為 NTD" + Environment.NewLine +
                                                        "請先檢查【來源Excel-[Currency]欄位】重新上傳 或 連絡MIS", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                                        txterr.Text += Environment.NewLine +
                                                       DateTime.Now.ToString() + Environment.NewLine +
                                                       "來源 第" + (i + 1) + "筆 " + "，轉換失敗!!" + Environment.NewLine +
                                                       "幣別 不等於 TWD 無法轉換為 NTD" + Environment.NewLine +
                                                       "請先檢查【來源Excel-[Currency]欄位】重新上傳 或 連絡MIS" + Environment.NewLine +
                                                       "===========";

                                        txt_path.Text = "";
                                        dgv_cfipo.DataSource = null;
                                        tabCtl_data.SelectedIndex = 0;
                                        btn_toerp.Enabled = false;
                                        btn_toerp.BackColor = System.Drawing.SystemColors.Control;
                                        btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
                                        dgv_excel.CurrentCell = dgv_excel.Rows[i].Cells["Currency"];

                                        //dt_訂單.Rows[i].RowState;

                                        return;
                                    }
                                    break;

                                //線別
                                case 5:
                                    str_sql_value = this.get_sql_value(data_type, dt_訂單.Rows[i][str_sql_column].ToString().Trim());
                                    break;

                                //客戶單價 Shipment Amount
                                case 11:
                                    // 假設 str_sql_column 和 data_type 已經定義並初始化
                                    string rawValue = dt_訂單.Rows[i][str_sql_column].ToString().Trim();
                                    str_sql_value = string.IsNullOrEmpty(rawValue) ? "'0'" :
                                                    this.get_sql_value(data_type, rawValue);
                                    break;
                                //客戶需求日期
                                case 15:
                                    str_sql_value = Convert.ToDateTime(dt_訂單.Rows[i][str_sql_column]).ToString("yyyyMMdd");//格式转换

                                    DateTime needdate = DateTime.ParseExact(str_sql_value, "yyyyMMdd", null, System.Globalization.DateTimeStyles.AllowWhiteSpaces);
                                    DateTime keydate = DateTime.ParseExact(textBox_單據日期.Text.ToString(), "yyyyMMdd", null, System.Globalization.DateTimeStyles.AllowWhiteSpaces);

                                    //判別 單據日期 大於 預交日期 則中斷，並關閉 轉換ERP格式按鈕
                                    if (keydate > needdate)
                                    {
                                        lab_status.Text = " 警告：請檢查【來源檔案-[Need By Date]】重新上傳!!";
                                        MessageBox.Show("來源 第" + (i + 1) + "筆 " + "，轉換失敗!!" + Environment.NewLine +
                                                        "【單據日期】大於【預交日期】" + Environment.NewLine +
                                                        "請先檢查【來源Excel-[Need By Date]欄位】重新上傳 或 連絡MIS", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                                        txterr.Text += Environment.NewLine +
                                                       DateTime.Now.ToString() + Environment.NewLine +
                                                       "來源 第" + (i + 1) + "筆 " + "，轉換失敗!!" + Environment.NewLine +
                                                        "【預交日期】大於【單據日期】" + Environment.NewLine +
                                                       "請先檢查【來源Excel-[Need By Date]欄位】重新上傳 或 連絡MIS" + Environment.NewLine +
                                                       "===========";

                                        txt_path.Text = "";
                                        dgv_cfipo.DataSource = null;
                                        tabCtl_data.SelectedIndex = 0;
                                        btn_toerp.Enabled = false;
                                        btn_toerp.BackColor = System.Drawing.SystemColors.Control;
                                        btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
                                        dgv_excel.CurrentCell = dgv_excel.Rows[i].Cells["Need By Date"];

                                        return;
                                    }
                                    break;

                                //備註
                                case 16:
                                    if (dt_訂單.Columns.Count == 13)
                                    {
                                        //最多13欄，判別最後一欄不是日期，
                                        if (dt_訂單.Rows[i][12].GetType() != typeof(DateTime))
                                        {
                                            str_sql_value = this.get_sql_value(data_type, dt_訂單.Rows[i][12].ToString().Trim());
                                        }
                                        else
                                        {
                                            str_sql_value = "''";
                                        }
                                    }
                                    //最後一欄有備註
                                    else if (dt_訂單.Columns.Count == 12 && dt_訂單.Rows[i][11].GetType() != typeof(DateTime))
                                    {
                                        str_sql_value = this.get_sql_value(data_type, dt_訂單.Rows[i][11].ToString().Trim());
                                    }
                                    else
                                    {
                                        str_sql_value = "''";
                                    }
                                    break;

                                default:
                                    str_sql_value = this.get_sql_value(data_type, dt_訂單.Rows[i][str_sql_column].ToString().Trim());
                                    
                                    break;
                            }
                            str_sql_columns = str_sql_columns + "[" + str_sql_column + "]" + ",";
                            str_sql_values = str_sql_values + str_sql_value + ",";
                        }

                        //新增至 CFIPO
                        //刪除最後的","符號
                        str_sql_columns = str_sql_columns.TrimEnd(new char[] { ',' });
                        str_sql_values = str_sql_values.TrimEnd(new char[] { ',' });
                        str_sql =
                            "insert into CFIPO (" + str_sql_columns + ")" + str_enter +
                            "VALUES(" + str_sql_values + ")";

                        this.to_ExecuteNonQuery(str_sql);
                        //MyCode.sqlExecuteNonQuery(str_sql);
                        str_sql_columns = "";
                        str_sql_values = "";
                    }
                    //TODO:將 轉換後的格式輸出
                    DataTable dt_cfipo = new DataTable();
                    this.sqlDataAdapter1.SelectCommand.CommandText = "select * from CFIPO order by ERP_Num";
                    this.sqlDataAdapter1.Fill(dt_cfipo);
                    dgv_cfipo.DataSource = dt_cfipo;

                    //str_sql = "select * from CFIPO order by ERP_Num";
                    //MyCode.Sql_dgv(str_sql, dt_cfipo, dgv_cfipo);

                    lab_status.Text = " 匯入 整理格式 完成";
                    tabCtl_data.SelectedIndex = 1;
                    scope.Complete();

                    txterr.Text += Environment.NewLine +
                                    DateTime.Now.ToString() + Environment.NewLine +
                                   ">> 匯入 整理格式 完成" + Environment.NewLine +
                                   "===========";

                }

                catch (Exception ex)
                {
                    lab_status.Text = " 錯誤：請檢查【來源檔案】重新上傳!!";
                    txt_path.Text = "";
                    MessageBox.Show("來源 第" + (i + 1) + "筆 " + "，轉換失敗!!" + Environment.NewLine +
                                    "【 " + ex.Message + " 】" + Environment.NewLine +
                                    "請先檢查【來源Excel格式】重新上傳 或 連絡MIS", "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txterr.Text += Environment.NewLine +
                                    DateTime.Now.ToString() + Environment.NewLine +
                                   "來源 第" + (i + 1) + "筆 " + "，轉換失敗!!" + Environment.NewLine +
                                    "【 " + ex.Message + " 】" + Environment.NewLine +
                                   "請先檢查【來源Excel格式】重新上傳 或 連絡MIS" + Environment.NewLine +
                                   "===========";
                    //關閉 ERP格式轉換 按鈕
                    tabCtl_data.SelectedIndex = 0;
                    btn_toerp.Enabled = false;
                    btn_toerp.BackColor = System.Drawing.SystemColors.Control;
                    btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
                    //關閉 上傳ERP按鈕
                    btn_erpup.Enabled = false;
                    btn_erpup.BackColor = System.Drawing.SystemColors.Control;
                    btn_erpup.ForeColor = System.Drawing.SystemColors.ControlText;
                    dgv_excel.CurrentCell = dgv_excel.Rows[i].Cells[1];


                    return;
                }
                //發生例外時，會自動rollback
                finally
                {
                    this.sqlConnection1.Close();
                }
            }

            //20240222 生管林玲禎 提出 轉換格式時，檢查EXCEL金額[Shipment Amount]是否與ERP相符但不卡控[轉換ERP格式]僅[緊示]
            //20240305 生管林玲禎 提出 欄位比對錯誤(單價-TD011)，正確為金額[TD012]，(COPMB.MB008 * CFIPO.Quantity) as TD012
            //TODO: 取得 COPTC.COPTD 客戶訂單單頭單身資料檔的欄位資料型態
            DataTable dt_CheckCFIPO_ShipmentAmount = new DataTable();

            string str_CheckCFIPO_ShipmentAmount = @"
    SELECT 
        CFIPO.[Number] as 'Number',
        CFIPO.Item as 'Item',
        INVMB.MB001 as 'TD004',
        (COPMB.MB008 * CFIPO.Quantity) as 'TD012',
        CFIPO.[Shipment Amount] as 'Shipment Amount'
    FROM 
        CFIPO
    LEFT JOIN 
        INVMB ON (SELECT MG002 FROM COPMG WHERE MG003 = CFIPO.Item AND MG001 = CFIPO.ERP_客代) = INVMB.MB001
    LEFT JOIN 
        (SELECT MB001, MB002, MB003, MB004, MB008, MB017 FROM COPMB a WHERE MB017 = (SELECT MAX(MB017) FROM COPMB WHERE MB002 = a.MB002 AND MB001 = a.MB001)) COPMB
    ON 
        (SELECT MG002 FROM COPMG WHERE MG003 = CFIPO.Item AND MG001 = CFIPO.ERP_客代) = COPMB.MB002 
        AND CFIPO.ERP_客代 = COPMB.MB001 
        AND CFIPO.UOM = COPMB.MB003 
        AND CFIPO.ERP_幣別 = COPMB.MB004";
            this.sqlDataAdapter1.SelectCommand.CommandText = str_CheckCFIPO_ShipmentAmount;
            this.sqlDataAdapter1.Fill(dt_CheckCFIPO_ShipmentAmount);

            // 假設 dt_CheckCFIPO_ShipmentAmount 是 DataTable 對象
            List<string> mismatchedRows = new List<string>();

            foreach (DataRow row in dt_CheckCFIPO_ShipmentAmount.Rows)
            {
                decimal shipmentAmount = Convert.ToDecimal(row["Shipment Amount"]);
                decimal td012 = Convert.ToDecimal(row["TD012"]);

                if (shipmentAmount != td012)
                {
                    // 如果值不同，收集信息
                    string message = $"Number: {row["Number"].ToString().Trim()}, " + Environment.NewLine +
                                     $"Item: {row["Item"].ToString().Trim()}, " + Environment.NewLine +
                                     $"ERP品號: {row["TD004"].ToString().Trim()}, " + Environment.NewLine +
                                     $"ERP單價: {Convert.ToInt32(row["TD012"]).ToString().Trim()}, " + Environment.NewLine +
                                     $"Shipment Amount: {row["Shipment Amount"].ToString().Trim()}";
                    mismatchedRows.Add(message);
                }
            }

            // 顯示所有不匹配的行信息
            if (mismatchedRows.Count > 0)
            {
                // 顯示警示
                MessageBox.Show("部分品項ERP單價與來源不符，共"+ mismatchedRows.Count.ToString() + "筆，請查看警示訊息", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txterr.Text += Environment.NewLine +
                          DateTime.Now.ToString() + Environment.NewLine +
                          "部分品項ERP單價與來源不符，共" + mismatchedRows.Count.ToString() + "筆，請查看警示訊息" + Environment.NewLine +
                          "===========";

                // 在 txterr 文本框中一起顯示
                foreach (string message in mismatchedRows)
                {
                    txterr.Text += Environment.NewLine +
                           DateTime.Now.ToString() + Environment.NewLine +
                           message + Environment.NewLine +
                           "===========";
                }
            }

            //匯入 cfipo 成功，開啟 ERP轉換格式 按鈕
            if (dgv_cfipo.DataSource != null)
            {
                btn_toerp.Enabled = true;
                btn_toerp.BackColor = System.Drawing.Color.SeaGreen;
                btn_toerp.ForeColor = System.Drawing.Color.White;
            }

        }

        private ListSortDirection sortdirection = ListSortDirection.Ascending;

        private DataGridViewColumn sortcolumn = null;

        private int sortColindex = -1;

        private void dgv_Sorted(object sender, EventArgs e)
        {

            sortcolumn = dgv_excel.SortedColumn;
            sortColindex = sortcolumn.Index;
            sortdirection =
            dgv_excel.SortOrder == SortOrder.Ascending ?
            ListSortDirection.Ascending : ListSortDirection.Descending;

        }

        //TODO:轉換 ERP格式 按鈕作業
        private void btn_toerp_Click(object sender, EventArgs e)
        {
            //判別當月第一筆
            if (str_key_客訂單號.Length >= 7 && str_key_客訂單號.Substring(4, 3).ToString() == "001")
            {
                str_key_客訂單號 = str_key_客訂單號.Substring(0, 4).ToString() + "000";
            }
            else
            {
                str_key_客訂單號 = (Convert.ToInt32(lab_num2.Text.ToString()) - 1).ToString().PadLeft(7, '0');
            }
            //單身 ERP格式
            //20210111 CONVERT(varchar(6) 改為 CONVERT(varchar(7)
            //", REPLICATE('0', (7 - LEN(CONVERT(varchar(6), (CFIPO.ERP_Num + " + str_key_客訂單號 + "))))) +CONVERT(varchar(6), (CFIPO.ERP_Num + " + str_key_客訂單號 + ")) as TD002" + str_enter +
            //20210913 升級GP4單身增加兩個欄位，計價數量(TD076).計價單位(TD077)，同原欄位 數量(TD008).單位(TD010)
            //,CFIPO.Quantity as TD076,CFIPO.UOM as TD077,
            DataTable dt_單身 = new DataTable();
            string str_sql_td =
            "SELECT '220' as TD001" + str_enter +
            ", REPLICATE('0', (7 - LEN(CONVERT(varchar(7), (CFIPO.ERP_Num + " + str_key_客訂單號 + "))))) +CONVERT(varchar(7), (CFIPO.ERP_Num + " + str_key_客訂單號 + ")) as TD002" + str_enter +
            ",CFIPO.ERP_序號 as TD003,INVMB.MB001 as TD004,INVMB.MB002 as TD005,INVMB.MB003 as TD006" + str_enter +
            ",INVMB.MB017 as TD007,CFIPO.Quantity as TD008,'0' as TD009,CFIPO.UOM as TD010" + str_enter +
            ",COPMB.MB008 as TD011,(COPMB.MB008 * CFIPO.Quantity) as TD012,CFIPO.[Need By Date] as TD013" + str_enter +
            ",CFIPO.Item as TD014,'' as TD015,'N' as TD016,'' as TD017,'' as TD018,'' as TD019" + str_enter +
            ",CFIPO.備註 as TD020,'N' as TD021,'0' as TD022,'' as TD023,'0' as TD024,'0' as TD025,'1' as TD026" + str_enter +
            ",'' as TD027,'' as TD028,'' as TD029,'0' as TD030,'0' as TD031,(CFIPO.Quantity / INVMD.MD004) as TD032" + str_enter +
            ",'0' as TD033,'0' as TD034,'0' as TD035,INVMB.MB090 as TD036,'' as TD037,'' as TD038" + str_enter +
            ",'' as TD039,'' as TD040,'' as TD041,'0' as TD042,'' as TD043,'' as TD044,'9' as TD045,'' as TD046" + str_enter +
            ",CFIPO.[Need By Date] as TD047,CFIPO.[Need By Date] as TD048,'1' as TD049,'0' as TD050" + str_enter +
            ",'0' as TD051,'0' as TD052,'0' as TD053,'0' as TD054,'0' as TD055,'' as TD056,'' as TD057" + str_enter +
            ",'' as TD058,'0' as TD059,'' as TD060,'0' as TD061,'' as TD062,'' as TD063,'' as TD064,'' as TD065" + str_enter +
            ",'' as TD066,'' as TD067,'' as TD068,'' as TD069" + str_enter +
            ",(select NN004 from CMSNN where NN001 = (select MA118 from COPMA where MA001 = CFIPO.ERP_客代)) as TD070" + str_enter +
            ",'' as TD071,'' as TD072,'' as TD073,'' as TD074,CFIPO.Quantity as TD076,CFIPO.UOM as TD077,'' as TD500,'0' as TD501,'' as TD502,'' as TD503" + str_enter +
            ",'' as TD504,'' as TD200,CFIPO.[Need By Date] as TD201,'0' as TD202,'' as TD203,'Y' as TD204,'' as TD205" + str_enter +
            "FROM CFIPO" + str_enter +
            "left join INVMB on(select MG002 from COPMG where MG003 = CFIPO.Item and MG001 = CFIPO.ERP_客代) = INVMB.MB001" + str_enter +
            "left join INVMD on(select MG002 from COPMG where MG003 = CFIPO.Item and MG001 = CFIPO.ERP_客代) = INVMD.MD001" + str_enter +
            "left join(select MB001, MB002, MB003, MB004, MB008, MB017 from COPMB a where MB017 = (select MAX(MB017) from COPMB where MB002 = a.MB002 and MB001 = a.MB001)) COPMB" + str_enter +
            "on(select MG002 from COPMG where MG003 = CFIPO.Item and MG001 = CFIPO.ERP_客代) = COPMB.MB002 and CFIPO.ERP_客代 = COPMB.MB001 and CFIPO.UOM = COPMB.MB003 and CFIPO.ERP_幣別 = COPMB.MB004" + str_enter +
            "order by TD002";

            this.sqlDataAdapter1.SelectCommand.CommandText = str_sql_td;
            this.sqlDataAdapter1.Fill(dt_單身);
            //MyCode.Sql_dt(str_sql_td, dt_單身);

            this.get_total(dt_單身);
        }
        //TODO:計算 單頭 總金額.總數量.總包裝數
        private void get_total(DataTable dt)
        {
            string str_sql_column_c = "", str_sql_value_c = "", str_sql_columns_c = "", str_sql_values_c = "";
            string str_sql_column_d = "", str_sql_value_d = "", str_sql_columns_d = "", str_sql_values_d = "";
            string data_type_d = "", data_type_c = "";
            bool bol_to_insert = false;
            this.ft_sum採購金額 = 0; this.ft_sum數量合計 = 0; this.ft_sum包裝數量合計 = 0;

            // 準備ERP上傳 字串
            str_sql_coptc = "";
            str_sql_coptd = "";

            //TODO: 取得 COPTC.COPTD 客戶訂單單頭單身資料檔的欄位資料型態
            DataTable dt_schema_c = new DataTable();
            DataTable dt_schema_d = new DataTable();

            string str_sqlschema_c =
                "select COLUMN_NAME,DATA_TYPE,IS_NULLABLE" + str_enter +
                "from INFORMATION_SCHEMA.COLUMNS" + str_enter +
                "where TABLE_NAME='COPTC'";
            this.sqlDataAdapter1.SelectCommand.CommandText = str_sqlschema_c;
            this.sqlDataAdapter1.Fill(dt_schema_c);
            //MyCode.Sql_dt(str_sqlschema_c, dt_schema_c);

            string str_sqlschema_d =
                "select COLUMN_NAME,DATA_TYPE,IS_NULLABLE" + str_enter +
                "from INFORMATION_SCHEMA.COLUMNS" + str_enter +
                "where TABLE_NAME='COPTD'";
            this.sqlDataAdapter1.SelectCommand.CommandText = str_sqlschema_d;
            this.sqlDataAdapter1.Fill(dt_schema_d);
            //MyCode.Sql_dt(str_sqlschema_d, dt_schema_d);

            //TODO: 填入[COMPANY],[CREATOR],[USR_GROUP] ,[CREATE_DATE] ,[MODIFIER],[MODI_DATE] ,[FLAG]
            string[] str_basic =
                {
                    this.str_廠別,
                    this.str_建立者ID,
                    this.str_建立者GP,
                    this.str_建立日期,
                    "",
                    "",
                    "0"
                };
            //TODO: 交易機制-單頭及單身寫入
            using (TransactionScope scope1 = new TransactionScope())
            {
                try
                {
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
                                        if (str_sql_value_d == "")
                                        {
                                            switch (str_sql_column_d)
                                            {
                                                case "TD011":
                                                    errtable = "單價";
                                                    break;
                                                default:
                                                    errtable = "未設定";
                                                    break;
                                            }

                                            lab_status.Text = " 錯誤：請檢查檔案重新上傳!!";
                                            MessageBox.Show("第" + (i + 1) + "筆 " + "，轉換ERP格式 失敗!!" + Environment.NewLine +
                                                            str_sql_column_d + " 欄位，【" + errtable + "】為空值" + Environment.NewLine +
                                                            "請先檢查【來源Excel檔案】及ERP系統確認 或 連絡MIS", "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                            txterr.Text += Environment.NewLine +
                                                            DateTime.Now.ToString() + Environment.NewLine +
                                                           "來源 第" + (i + 1) + "筆 " + "，轉換ERP格式 失敗!!" + Environment.NewLine +
                                                           str_sql_column_d + " 欄位，【" + errtable + "】為空值" + Environment.NewLine +
                                                           "請先檢查【來源Excel檔案】及ERP系統確認 或 連絡MIS" + Environment.NewLine +
                                                           "===========";

                                            btn_toerp.Enabled = false;
                                            btn_toerp.BackColor = System.Drawing.SystemColors.Control;
                                            btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
                                            dgv_excel.CurrentCell = dgv_excel.Rows[i].Cells[1];

                                            return;
                                        }
                                        else
                                        {

                                            bol_to_insert = true;
                                        }
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
                        str_sql_d =
                            "insert into COPTD (" + str_sql_columns_d + ")" + str_enter +
                            "VALUES(" + str_sql_values_d + ")";

                        // 上傳ERP單身字串整理後，屆時一次上傳
                        str_sql_coptd += str_sql_d + str_enter;

                        this.to_ExecuteNonQuery(str_sql_d);
                        //MyCode.sqlExecuteNonQuery(str_sql_d);

                        //加總
                        ft_sum採購金額 += Convert.ToDouble(dt.Rows[i]["TD012"].ToString());
                        ft_sum數量合計 += Convert.ToDouble(dt.Rows[i]["TD008"].ToString()); ;
                        ft_sum包裝數量合計 += Convert.ToDouble(dt.Rows[i]["TD032"].ToString()); ;

                        //TODO:判別 單身單號與下一筆不同 則新增 單頭
                        //20210111 CONVERT(varchar(6) 改為 CONVERT(varchar(7)
                        // ", REPLICATE('0', (7 - LEN(CONVERT(varchar(6), (CFIPO.ERP_Num + " + str_key_客訂單號 + "))))) +CONVERT(varchar(6), (CFIPO.ERP_Num + " + str_key_客訂單號 + ")) as TC002" + str_enter +
                        //20240229 生管 林玲禎提出，財務聯絡因財務報表製作時需要［部門別］資訊 COPMA.MA015對應訂單單頭COPTC.TC005
                        if ((i != dt.Rows.Count - 1 && (dt.Rows[i]["TD002"].ToString() != dt.Rows[i + 1]["TD002"].ToString())) || i == dt.Rows.Count - 1)
                        {
                            DataTable dt_單頭 = new DataTable();
                            string str_sql_tc =
                            "select * from (select'220' as TC001" + str_enter +
                            ", REPLICATE('0', (7 - LEN(CONVERT(varchar(7), (CFIPO.ERP_Num + " + str_key_客訂單號 + "))))) +CONVERT(varchar(7), (CFIPO.ERP_Num + " + str_key_客訂單號 + ")) as TC002" + str_enter +
                            ",'" + textBox_單據日期.Text.ToString().Trim() + "' as TC003 ,CFIPO.ERP_客代 as TC004,COPMA.MA015 as TC005" + str_enter +
                            ",CFIPO.線別 as TC006,'002' as TC007,COPMA.MA014 as TC008 " + str_enter +
                            ",(select MG004 from CMSMG where MG001 = COPMA.MA014 and MG002 = (select MAX(MG002) from CMSMG where MG001 = COPMA.MA014))  as TC009" + str_enter +
                            ",(COPMA.MA080 + ' ' +COPMA.MA027) as TC010,COPMA.MA064 as TC011,CFIPO.ERP_客單 as TC012,COPMA.MA030 as TC013,COPMA.MA031 as TC014" + str_enter +
                            ",'' as TC015,COPMA.MA038 as TC016,'' as TC017,COPMA.MA005 as TC018,COPMA.MA048 as TC019,'' as TC020" + str_enter +
                            ",'' as TC021,COPMA.MA056 as TC022,COPMA.MA057 as TC023,COPMA.MA058 as TC024,'' as TC025,COPMA.MA059 as TC026" + str_enter +
                            ",'N' as TC027,'0' as TC028,'" + ft_sum採購金額 + "' as TC029,'0' as TC030,'" + ft_sum數量合計 + "' as TC031" + str_enter +
                            ",CFIPO.ERP_客代 as TC032,'' as TC033,'' as TC034,COPMA.MA051 as TC035,'' as TC036,'' as TC037,'' as TC038" + str_enter +
                            ",'" + textBox_單據日期.Text.ToString().Trim() + "' as TC039,'' as TC040" + str_enter +
                            ",(select NN004 from CMSNN where NN001 = (select MA118 from COPMA where MA001 = CFIPO.ERP_客代))  as TC041" + str_enter +
                            ",COPMA.MA083 as TC042,'0' as TC043,'0' as TC044,COPMA.MA095 as TC045,'" + ft_sum包裝數量合計 + "' as TC046" + str_enter +
                            ",'' as TC047,'N' as TC048,'' as TC049,'N' as TC050,'' as TC051,'0' as TC052,COPMA.MA003 as TC053" + str_enter +
                            ",'' as TC054,'' as TC055,'1' as TC056,'N' as TC057,'' as TC058,'' as TC059,'N' as TC060,'' as TC061" + str_enter +
                            ",'' as TC062,(COPMA.MA079 + ' ' + COPMA.MA025) as TC063,COPMA.MA026 as TC064,COPMA.MA003 as TC065,COPMA.MA006 as TC066" + str_enter +
                            ",COPMA.MA008 as TC067,'1' as TC068,'0000' as TC069,'N' as TC070,COPMA.MA110 as TC071,'0' as TC072" + str_enter +
                            ",'0' as TC073,'' as TC074,'' as TC075,'' as TC076,'N' as TC077,COPMA.MA118 as TC078,'' as TC079" + str_enter +
                            ",'' as TC080,'' as TC081,COPMA.MA076 as TC082,COPMA.MA077 as TC083,COPMA.MA078 as TC084,'' as TC085" + str_enter +
                            ",'' as TC086,'' as TC087,'' as TC088,'' as TC089,'' as TC090,COPMA.MA123 as TC091,'' as TC092,'' as TC200" + str_enter +
                            "from CFIPO" + str_enter +
                            "left join COPMA on CFIPO.ERP_客代 = COPMA.MA001" + str_enter +
                            "where CFIPO.ERP_序號 = '0001')CFITC where TC002 ='" + dt.Rows[i]["TD002"].ToString().Trim() + "'";

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
                                          , str_建立者ID, str_建立日期, dt_單頭.Rows[x]["TC001"], dt_單頭.Rows[x]["TC002"], "COPTC", "客戶訂單 匯入", "新增客戶訂單單頭", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

                                // 上傳ERP Log字串整理後，屆時一次上傳
                                str_sql_logs += str_sql_log + str_enter;

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
                        "where TD001 ='220' and TD002 between '" + dt.Rows[0]["TD002"].ToString() + "' and '" + dt.Rows[i - 1]["TD002"].ToString() + "' order by TD002";
                    //str_sql =
                    //    "select TD001 as 單別 ,TD002 as 單號 ,TD003 as 序號 ,TD004 as 品號 ,TD005 as 品名,TD006 as 規格 " +
                    //    ",TD007 as 庫別 ,TD010 as 單位 ,TD011 as 單價 ,TD008 as 訂單數量,TD012 as 金額 ,TD036 as 包裝單位 " +
                    //    ",TD032 as 訂單包裝數量 ,TD014 as 客戶品號,TD020 as 備註 ,TD202 as 實際可交貨數,TD013 as 預交日 " +
                    //    ",TD047 as 原預交日 ,TD048 as 排定交貨日 ,TD201 as 希望交貨日 from COPTD " +
                    //    "where TD001 ='220' and TD002 between '" + dt.Rows[0]["TD002"].ToString() + "' and '" + dt.Rows[i - 1]["TD002"].ToString() + "' order by TD002";

                    this.sqlDataAdapter1.Fill(dt_coptd);
                    dgv_td.DataSource = dt_coptd;
                    //MyCode.Sql_dgv(str_sql, dt_coptd, dgv_td);

                    DataTable dt_coptc = new DataTable();
                    this.sqlDataAdapter1.SelectCommand.CommandText =
                        "select TC001 as 單別,TC002 as 單號,TC003 as 訂單日期,TC004 as 客戶代號,TC005 as 部門代號" +
                        ",TC006 as 業務人員,TC008 as 交易幣別,TC009 as 匯率,TC041 as 營業稅率,TC012 as 客戶單號" +
                        ",TC029 as 訂單金額,TC030 as 訂單稅額,TC031 as 總數量,TC046 as 總包裝數量,TC053 as 客戶全名" +
                        ",TC010 as 送貨地址_一,TC014 as 付款條件,TC018 as 連絡人 from COPTC " +
                        "where TC001 ='220' and TC002 between '" + dt.Rows[0]["TD002"].ToString() + "' and '" + dt.Rows[i - 1]["TD002"].ToString() + "' order by TC002";
                    //str_sql =
                    //    "select TC001 as 單別,TC002 as 單號,TC003 as 訂單日期,TC004 as 客戶代號,TC005 as 部門代號" +
                    //    ",TC006 as 業務人員,TC008 as 交易幣別,TC009 as 匯率,TC041 as 營業稅率,TC012 as 客戶單號" +
                    //    ",TC029 as 訂單金額,TC030 as 訂單稅額,TC031 as 總數量,TC046 as 總包裝數量,TC053 as 客戶全名" +
                    //    ",TC010 as 送貨地址_一,TC014 as 付款條件,TC018 as 連絡人 from COPTC " +
                    //    "where TC001 ='220' and TC002 between '" + dt.Rows[0]["TD002"].ToString() + "' and '" + dt.Rows[i - 1]["TD002"].ToString() + "' order by TC002";
                    this.sqlDataAdapter1.Fill(dt_coptc);
                    dgv_tc.DataSource = dt_coptc;
                    //MyCode.Sql_dgv(str_sql, dt_coptc, dgv_tc);

                    lab_status.Text = " ERP格式轉換完成";
                    tabCtl_data.SelectedIndex = 3;
                    tabCtl_data.SelectedIndex = 2;

                    //確認資料轉換成 ERP格式後，開啟 上傳ERP按鈕
                    btn_erpup.Enabled = true;
                    btn_erpup.BackColor = System.Drawing.Color.SteelBlue;
                    btn_erpup.ForeColor = System.Drawing.Color.White;

                    txterr.Text += Environment.NewLine +
                                    DateTime.Now.ToString() + Environment.NewLine +
                                   ">> ERP格式轉換完成" + Environment.NewLine +
                                   "===========";

                    //scope.Complete();
                }
                catch (Exception ex)
                {
                    lab_status.Text = " 錯誤：請檢查檔案重新上傳!!";
                    MessageBox.Show("第" + (i + 1) + "筆 " + "，轉換ERP格式 失敗!!" + Environment.NewLine +
                                    "【 " + ex.Message + " 】" + Environment.NewLine +
                                    "請先檢查【來源Excel格式】重新上傳 或 連絡MIS", "錯誤訊息", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    txterr.Text += Environment.NewLine +
                                    DateTime.Now.ToString() + Environment.NewLine +
                                   "來源 第" + (i + 1) + "筆 " + "，轉換ERP格式 失敗!!" + Environment.NewLine +
                                    "【 " + ex.Message + " 】" + Environment.NewLine +
                                   "請先檢查【來源Excel格式】重新上傳 或 連絡MIS" + Environment.NewLine +
                                   "===========";

                    
                    btn_toerp.Enabled = false;
                    btn_toerp.BackColor = System.Drawing.SystemColors.Control;
                    btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
                    dgv_excel.CurrentCell = dgv_excel.Rows[i].Cells[1];

                    return;
                }
                //發生例外時，會自動rollback
                finally
                {
                    this.sqlConnection1.Close();
                }
            }
        }
        //TODO:上傳至 ERP系統 按鈕
        private void btn_erpup_Click(object sender, EventArgs e)
        {
            //TODO:交易機制-確認 資料無誤，上傳ERP系統
            using (TransactionScope scope = new TransactionScope())
            {
                try
                {
                    DialogResult Result = MessageBox.Show("請再次確認資料", "確認上傳ERP", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                    if (Result == DialogResult.OK)
                    {
                        this.to_ExecuteNonQuery(str_sql_coptc);
                        this.to_ExecuteNonQuery(str_sql_coptd);
                        this.to_ExecuteNonQuery(str_sql_logs);

                        //MyCode.sqlExecuteNonQuery(str_sql_coptc);
                        //MyCode.sqlExecuteNonQuery(str_sql_coptd);
                        scope.Complete();

                        MessageBox.Show("已上傳至ERP系統");

                        txterr.Text += Environment.NewLine +
                                    DateTime.Now.ToString() + Environment.NewLine +
                                   ">> 已上傳至ERP系統" + Environment.NewLine +
                                   "===========";
                        //TODO:上傳 ERP系統完成後，將單據號碼.單據日期.檔案路徑.EXCEL匯入.CFIPO畫面清除，
                        //並關閉 轉換ERP格式及上傳ERP按鈕
                        lab_num2.Text = "";
                        textBox_單據日期.Text = "";
                        txt_path.Text = "";
                        dgv_excel.DataSource = null;
                        dgv_cfipo.DataSource = null;

                        btn_toerp.Enabled = false;
                        btn_toerp.BackColor = System.Drawing.SystemColors.Control;
                        btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;

                        btn_erpup.Enabled = false;
                        btn_erpup.BackColor = System.Drawing.SystemColors.Control;
                        btn_erpup.ForeColor = System.Drawing.SystemColors.ControlText;

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
                                   "請先檢查【來源Excel格式】重新上傳 或 連絡MIS" + Environment.NewLine +
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

        bool IsToForm1 = false; //紀錄是否要回到Form1
        protected override void OnClosing(CancelEventArgs e) //在視窗關閉時觸發
        {
            //DialogResult dr = MessageBox.Show("\"是\"回到主畫面 \r\n \"否\"關閉程式", "是否要關閉程式"
            //    , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            //if (dr == DialogResult.Yes)
            ////{
            ////TOYOINK_dev.fm_menu fm_menu = new TOYOINK_dev.fm_menu();
            //IsToForm1 = true;
            ////}

            //base.OnClosing(e);
            //if (IsToForm1) //判斷是否要回到Form1
            //{
            //    this.DialogResult = DialogResult.Yes; //利用DialogResult傳遞訊息
            //    fm_menu fm_menu = (fm_menu)this.Owner; //取得父視窗的參考
            //    fm_menu.show_fmlogin_CheckForm(1);
            //}
            //else
            //{
            //    this.DialogResult = DialogResult.No;
            //}
            Environment.Exit(Environment.ExitCode);
        }

    }
}
/*
 * A01A 
 SELECT COLUMN_NAME AS 欄位名稱,
DATA_TYPE AS 欄位型態,
CHARACTER_MAXIMUM_LENGTH AS 長度限制,
IS_NULLABLE AS 是否允許空值, 
COLUMN_DEFAULT AS 預設值
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = N'CFIPO'

 * 欄位名稱	欄位型態	長度限制	是否允許空值	預設值
ERP_Num	int	NULL	YES	NULL
ERP_序號	nvarchar	4	YES	NULL
ERP_客代	nvarchar	50	YES	NULL
ERP_客單	nvarchar	50	YES	NULL
ERP_幣別	nvarchar	50	YES	NULL
線別	nvarchar	50	YES	NULL
Sample	nvarchar	50	YES	NULL
Number	nvarchar	50	YES	NULL
Item	nvarchar	50	YES	NULL
Item Description	nvarchar	100	YES	NULL
UOM	nvarchar	50	YES	NULL
Shipment Amount	decimal	NULL	YES	NULL
Quantity	decimal	NULL	YES	NULL
Supplier	nvarchar	50	YES	NULL
Currency	nvarchar	50	YES	NULL
Need By Date	nvarchar	50	YES	NULL
備註	nvarchar	100	YES	NULL
 */
