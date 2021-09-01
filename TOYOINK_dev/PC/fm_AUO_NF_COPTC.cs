using System;
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

    public partial class fm_AUO_NF_COPTC : Form
    {
        /*
         * 20210623 開發完成，生管林玲禎提出，參考 客戶訂單For 友達(fm_AUOCOPTC)，修改來源Excel判別
         * 20210901 生管林玲禎提出 TD201 希望交貨日改為空值【CFIPO.[Need By Date] as TD201】 ->【'' as TD201】，
         *          經了解，[希望交貨日]為個案欄位，目前已無使用，了解使用者需求，初判該欄位應可使用標準[排定交貨日]應用，
         *          個案書內容 W71565_016.W71565_018.W71565_022.W71565_064
         * 
         */
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
       
        public fm_AUO_NF_COPTC()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();

            //MyCode.strDbCon = MyCode.strDbConLeader;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConLeader;

            MyCode.strDbCon = MyCode.strDbConA01A;
            this.sqlConnection1.ConnectionString = MyCode.strDbConA01A;

            //MyCode.strDbCon = "packet size=4096;user id=pwuser;password=sqlmis003;data source=192.168.128.219;persist security info=False;initial catalog=A01A;";
            ////MyCode.strDbCon = "packet size=4096;user id=yj.chou;password=yjchou3369;data source=192.168.128.219;persist security info=False;initial catalog=Leader;";
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

        private void btn_NeedDate_Click(object sender, EventArgs e)
        {
            //TODO:單頭及單身若不為空值，表示已轉換ERP格式，需重新轉換 或 資料已上傳ERP，需重新選擇日期
            //資料上傳ERP後，dgv_excel會清空
            //if (dgv_tc.DataSource != null || dgv_td.DataSource != null || dgv_excel.DataSource != null)
            if (btn_toerp.Enabled == true || btn_erpup.Enabled == true)

            {
                DialogResult Result = MessageBox.Show("修改 需求日期 後，需重新【選擇檔案】", "Excel檔案已匯入", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

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
                               " 修改需求日期，請重新【選擇檔案】" + Environment.NewLine +
                               "===========";

                    this.fm_月曆 = new 月曆(this.txt_NeedDate, this.btn_NeedDate, "需求日期");
                }
                else if (Result == DialogResult.Cancel)
                {
                    return;
                }
            }
            else
            {
                this.fm_月曆 = new 月曆(this.txt_NeedDate, this.btn_NeedDate, "需求日期");
                btn_file.Enabled = true;
                lab_status.Text = "請 選擇檔案";

            }
        }
        private void txterr_TextChanged(object sender, EventArgs e)
        {
            txterr.SelectionStart = txterr.Text.Length;
            txterr.ScrollToCaret();  //跳到遊標處 
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

        private void fm_AUO_NF_COPTC_Load(object sender, EventArgs e)
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
            else if (txt_NeedDate.Text.Length == 0)
            {
                MessageBox.Show("需先選擇【需求日期】");
                return;
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

            dgv_excel.DataSource = MyClass.ReadExcelToTable("fm_AUO_NF_COPTC",txt_path.Text.ToString(),"1=1");

            DataTable dt_All_Line = new DataTable();
            dt_All_Line = (DataTable)this.dgv_excel.DataSource;

            //使用Linq進行查詢北廠線別
            string[] Array_NF = new string[] { "M01", "M02", "M11", "L7A", "M12", "L8B" };

            var Linq_NF = from r in dt_All_Line.AsEnumerable()
                              where Array_NF.Contains(r.Field<string>("Org"))
                              select r;
            DataTable dt_訂單 = new DataTable();
            dt_訂單 = Linq_NF.CopyToDataTable();

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
                                    else if ((dt_訂單.Rows[i]["PO NO"].ToString().Trim()) == (dt_訂單.Rows[i - 1]["PO NO"].ToString().Trim()))
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
                                    else if ((dt_訂單.Rows[i]["PO NO"].ToString().Trim()) == (dt_訂單.Rows[i - 1]["PO NO"].ToString().Trim()))
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
                                    if ((dt_訂單.Rows[i]["Org"].ToString().Trim()) == "M01")
                                    {
                                        str_sql_value = this.get_sql_value(data_type, "AU-TY");
                                    }
                                    else
                                    {
                                        str_sql_value = this.get_sql_value(data_type, "AU-TC");
                                    }
                                    break;

                                //ERP 客戶單號
                                case 3:
                                    //if ((dt_訂單.Rows[i][0].ToString().Trim()) == "C5E")
                                    //{
                                    //    //[線別]+'-'+[Number]+'-HC' 
                                    //    str_sql_value = this.get_sql_value(data_type, (dt_訂單.Rows[i][0].ToString().Trim() + '-' + dt_訂單.Rows[i][2].ToString().Trim() + "-HC"));
                                    //}
                                    //else
                                    //{
                                        //str_sql_value = this.get_sql_value(data_type, (dt_訂單.Rows[i][0].ToString().Trim() + '-' + dt_訂單.Rows[i][2].ToString().Trim() + "-LT"));
                                    //}

                                    str_sql_value = this.get_sql_value(data_type, (dt_訂單.Rows[i]["PO NO"].ToString().Trim()));

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
                                                        "請先檢查【來源Excel-[PO NO]欄位】重新上傳 或 連絡MIS", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                                        txterr.Text += Environment.NewLine +
                                                       DateTime.Now.ToString() + Environment.NewLine +
                                                       "來源 第" + (i + 1) + "筆 " + "，轉換失敗!!" + Environment.NewLine +
                                                        "與【單據：" + dt_TC012.Rows[0]["TC001"].ToString().Trim() + "-" + dt_TC012.Rows[0]["TC002"].ToString().Trim() + "】，" +
                                                       "【客戶單號：" + str_TC012 + "】重複!!" + Environment.NewLine +
                                                       "請先檢查【來源Excel-[PO NO]欄位】重新上傳 或 連絡MIS" + Environment.NewLine +
                                                       "===========";

                                        dgv_cfipo.DataSource = null;
                                        tabCtl_data.SelectedIndex = 0;
                                        btn_toerp.Enabled = false;
                                        btn_toerp.BackColor = System.Drawing.SystemColors.Control;
                                        btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
                                        dgv_excel.CurrentCell = dgv_excel.Rows[i].Cells["PO NO"];

                                        return;
                                    }
                                    break;

                                //ERP 幣別 Cur
                                case 4:
                                    if (dt_訂單.Rows[i]["Cur"].ToString().Trim() == "TWD")
                                    {
                                        str_sql_value = this.get_sql_value(data_type, "NTD");
                                    }
                                    else
                                    {
                                        lab_status.Text = " 警告：請檢查【來源檔案-[Cur]】重新上傳!!";
                                        MessageBox.Show("來源 第" + (i + 1) + "筆 " + "，轉換失敗!!" + Environment.NewLine +
                                                        "幣別 不等於 TWD 無法轉換為 NTD" + Environment.NewLine +
                                                        "請先檢查【來源Excel-[Cur]欄位】重新上傳 或 連絡MIS", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                                        txterr.Text += Environment.NewLine +
                                                       DateTime.Now.ToString() + Environment.NewLine +
                                                       "來源 第" + (i + 1) + "筆 " + "，轉換失敗!!" + Environment.NewLine +
                                                       "幣別 不等於 TWD 無法轉換為 NTD" + Environment.NewLine +
                                                       "請先檢查【來源Excel-[Cur]欄位】重新上傳 或 連絡MIS" + Environment.NewLine +
                                                       "===========";

                                        txt_path.Text = "";
                                        dgv_cfipo.DataSource = null;
                                        tabCtl_data.SelectedIndex = 0;
                                        btn_toerp.Enabled = false;
                                        btn_toerp.BackColor = System.Drawing.SystemColors.Control;
                                        btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
                                        dgv_excel.CurrentCell = dgv_excel.Rows[i].Cells["Cur"];

                                        //dt_訂單.Rows[i].RowState;

                                        return;
                                    }
                                    break;

                                //線別 / Location
                                case 5:
                                    switch (dt_訂單.Rows[i]["Location"].ToString().Trim()) 
                                    {
                                        case "M01":
                                        case "M01-L5A":
                                            str_sql_value = this.get_sql_value(data_type, "C5A-B");
                                            break;
                                        case "M02":
                                        case "M02-L6B":
                                            str_sql_value = this.get_sql_value(data_type, "C6B");
                                            break;
                                        case "M11":
                                        case "M11-L6A":
                                        case "M11-C6A":
                                            str_sql_value = this.get_sql_value(data_type, "C6A");
                                            break;
                                        case "L7A-L5C":
                                        case "L7A-C5C":
                                            str_sql_value = this.get_sql_value(data_type, "C5C");
                                            break;
                                        case "L7A":
                                            str_sql_value = this.get_sql_value(data_type, "C7A");
                                            break;
                                        case "M12":
                                        case "M12-L7B":
                                            str_sql_value = this.get_sql_value(data_type, "C7B");
                                            break;
                                        case "M12-L8A":
                                            str_sql_value = this.get_sql_value(data_type, "C8A");
                                            break;
                                        case "L8B":
                                            str_sql_value = this.get_sql_value(data_type, "C8B");
                                            break;
                                        default:
                                            break;
                                    }
                                    break;

                                //Number 客戶單號
                                case 7:
                                    str_sql_value = this.get_sql_value(data_type, (dt_訂單.Rows[i]["PO NO"].ToString().Trim()));
                                    break;
                                //Item 客戶品號
                                case 8:
                                    str_sql_value = this.get_sql_value(data_type, (dt_訂單.Rows[i]["Item No"].ToString().Trim()));
                                    break;
                                //Item Description 客戶品號名稱
                                case 9:
                                    str_sql_value = this.get_sql_value(data_type, (dt_訂單.Rows[i]["Item Desc)"].ToString().Trim()));
                                    break;
                                //UOM 重量
                                case 10:
                                    str_sql_value = this.get_sql_value(data_type, "KG");
                                    break;
                                //Quantity
                                case 12:
                                    str_sql_value = this.get_sql_value(data_type, (dt_訂單.Rows[i]["Qty(Order Quantity)"].ToString().Trim()));
                                    break;
                                //Currency 幣別
                                case 14:
                                    str_sql_value = this.get_sql_value(data_type, (dt_訂單.Rows[i]["Cur"].ToString().Trim()));
                                    break;
                                //客戶需求日期
                                case 15:
                                    str_sql_value = this.get_sql_value(data_type, txt_NeedDate.Text.ToString().Trim());

                                    DateTime needdate = DateTime.ParseExact(txt_NeedDate.Text.ToString().Trim(), "yyyyMMdd", null, System.Globalization.DateTimeStyles.AllowWhiteSpaces);
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

                                        return;
                                    }
                                    break;

                                //Sample
                                case 6:
                                //Supplier
                                case 13:
                                //備註
                                case 16:
                                    str_sql_value = "''";
                                    break;

                                //Shipment Amount
                                case 11:
                                    str_sql_value = this.get_sql_value(data_type, "0") ;
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
            //20210901 生管林玲禎提出 TD201 希望交貨日改為空值【CFIPO.[Need By Date] as TD201】 ->【'' as TD201】
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
            ",'' as TD071,'' as TD072,'' as TD073,'' as TD074,'' as TD500,'0' as TD501,'' as TD502,'' as TD503" + str_enter +
            ",'' as TD504,'' as TD200,'' as TD201,'0' as TD202,'' as TD203,'Y' as TD204,'' as TD205" + str_enter +
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
                        if ((i != dt.Rows.Count - 1 && (dt.Rows[i]["TD002"].ToString() != dt.Rows[i + 1]["TD002"].ToString())) || i == dt.Rows.Count - 1)
                        {
                            DataTable dt_單頭 = new DataTable();
                            string str_sql_tc =
                            "select * from (select'220' as TC001" + str_enter +
                            ", REPLICATE('0', (7 - LEN(CONVERT(varchar(7), (CFIPO.ERP_Num + " + str_key_客訂單號 + "))))) +CONVERT(varchar(7), (CFIPO.ERP_Num + " + str_key_客訂單號 + ")) as TC002" + str_enter +
                            ",'" + textBox_單據日期.Text.ToString().Trim() + "' as TC003 ,CFIPO.ERP_客代 as TC004,'' as TC005" + str_enter +
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
                                          , str_建立者ID, str_建立日期, dt_單頭.Rows[x]["TC001"], dt_單頭.Rows[x]["TC002"], "COPTC", "fm_AUO_NF_COPTC", "新增客戶訂單單頭", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

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
