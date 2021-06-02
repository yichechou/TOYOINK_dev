using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Myclass;
using ClosedXML.Excel;
using System.Globalization;
using System.IO;
using System.Transactions;

namespace TOYOINK_dev
{
    /*
     * 20210315 cond_y 與 cond_N 加入【or (MN001 in ('2218','2219','2166') and MN018 <> 0)】
     * 20210315 sql_Details_Primary 與 sql_Quarter_Primary 更新備註欄位 從科目代號改為 依備註內單別[710T.C71T]判別
     * 20210413 遺漏科目3218 上期損益.7108 股利收入，修改判別條件原科目名稱判別(MA003 like '%關係%' or MA003 like '%TOPPAN%')，改為關係人代號 MA012 <> ''判別
     *  科目單月條件異動，加入科目7開頭，並分為left(MN001,1) in ('1','2','3','7') and MN020 = 'N' and MA012 <> '' and MN004 <= '20210331' 與
     *  left(MN001,1) in ('4','5','6') and MN020 = 'N' and MA012 <> '' and MN004 like '202103%'，其中3.7 改為累計方式呈現
     */
    public partial class fm_Acc_RelatedVOU : Form
    {
        public MyClass MyCode;
        月曆 fm_月曆;
        string save_as_RelatedVOU = "", temp_excel_RelatedVOU;
        string createday = DateTime.Now.ToString("yyyy/MM/dd");
        int opencode = 0,check_Prep = 0;

        string str_date_s, str_date_m_s, str_date_ym_s;
        string str_date_e, str_date_m_e, str_date_ym_e, str_date_y_e;

        string str_Prep_date_s, str_Prep_date_m_s, str_Prep_date_ym_s;
        string str_Prep_date_e, str_Prep_date_m_e, str_Prep_date_ym_e, str_Prep_date_y_e, str_Prep_date_LastY;
        string str_Create_date_ym;

        string cond_Cost_N, cond_Cost_y, cond_Posting, cond_y, cond_N, cond_LastY3456, cond_LastN3456;
        string sql_Details_Primary, sql_Quarter_Primary;
        string cond_Month_A,cond_Month_B, cond_Quarter;
        string cond_MonthSum, cond_QuartarSum;


        string defaultfilePath = "";

        DateTime date_s, date_e;
        DateTime Prep_date_s, Prep_date_e;


        DataTable dt_Posting = new DataTable();  //確認是否過帳
        DataTable dt_Cost_N = new DataTable();  //確認是否已產生成本結轉
        DataTable dt_Cost_y = new DataTable();  //確認是否已產生成本結轉
        DataTable dt_y = new DataTable();  //指定結案 y
        DataTable dt_N = new DataTable();  //取消結案 N
        DataTable dt_LastY3456_y = new DataTable();  //去年度科目3.4.5.6指定結案

        DataTable dt_ReportCheckCost = new DataTable();  //確認是否已產生成本結轉
        DataTable dt_ReportCheckPosting = new DataTable();  //確認是否過帳
        DataTable dt_MonthDetails = new DataTable("dt_MonthDetails");  //單月明細
        DataTable dt_QuarterDetails = new DataTable("dt_QuarterDetails");  //累計明細
        DataTable dt_MonthSum = new DataTable("dt_MonthSum");  //單月彙總
        DataTable dt_QuarterSum = new DataTable("dt_QuarterSum");  //累計彙總

        DataTable dt_修改者;


        string str_廠別 = "A01A", str_修改者ID = "", str_修改者GP = "", str_修改日期 = "";



        public fm_Acc_RelatedVOU()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();
            MyCode.strDbCon = "packet size=4096;user id=pwuser;password=sqlmis003;data source=192.168.128.219;persist security info=False;initial catalog=A01A;";
            //MyCode.strDbCon = "packet size=4096;user id=yj.chou;password=asdf0000;data source=192.168.128.219;persist security info=False;initial catalog=Leader;";
            temp_excel_RelatedVOU = @"\\192.168.128.219\Company\MIS自開發主檔\會計報表公版\關係人交易_temp.xlsx";
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

        private void fm_Acc_RelatedVOU_Load(object sender, EventArgs e)
        {
            //TODO:匯入ERP 可修改關係人立沖 使用者清單
            dt_修改者 = new DataTable();
            MyCode.Sql_dt("select MF001,MF001 + MF002 as 人員,MF002,MF004 from ADMMF", dt_修改者);

            this.cob_修改者.Items.Clear();
            string str_修改者 = "";
            int check = 0;

            for (int i = 0; i < dt_修改者.Rows.Count; i++)
            {
                str_修改者 = this.dt_修改者.Rows[i]["MF002"].ToString().Trim();
                this.cob_修改者.Items.Add(dt_修改者.Rows[i]["人員"].ToString().Trim());

                if (str_修改者 == loginName || (loginName == "周怡甄" && str_修改者 == "MIS用"))
                {
                    this.cob_修改者.SelectedIndex = i;
                    check = 1;
                }

            }
            if (check == 0)
            {
                MessageBox.Show("非會計人員不能使用", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txterr.Text += Environment.NewLine +
                           DateTime.Now.ToString() + Environment.NewLine +
                           "非會計人員不能使用" + Environment.NewLine +
                           "===========";
                btn_file.Enabled = false;
                cob_修改者.Enabled = false;

                //fm_login fm_login = new fm_login();

                //fm_login.Show();
                //this.Hide();
                return;
            }

            //TODO:格式化 修改日期
            lab_Nowdate.Text = DateTime.Now.ToString("yyyyMMdd");
            str_修改日期 = lab_Nowdate.Text.ToString().Trim();
            txt_Prep_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_Prep_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
            txt_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
            string filder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            txt_path.Text = filder;

            cond_Posting = @"TA010='Y' and TA001 not in ('915') and TA011 = 'N'";
            cond_Cost_N = @"MN020 = 'N' and (MA003 like '%關係%' or MA003 like '%TOPPAN%') and MN010 like '%成本結轉%'";
            cond_Cost_y = @"MN020 = 'y' and (MA003 like '%關係%' or MA003 like '%TOPPAN%') and MN010 like '%成本結轉%'";
            //20210315 cond_y 與 cond_N 加入【or(MN001 in ('2218', '2219', '2166') and MN018 <> 0)】
            cond_y = @"MN020 = 'N' and (MA003 like '%關係%' or MA003 like '%TOPPAN%')
                          and ( MN010 like '%迴轉%' or MN010 like '%評價%' or MN010 like '%手動結案%' or MN010 like '%成本結轉%'
                          or (MN001 in ('2218','2219','2166') and MN018 <> 0))";
            //cond_N = @"((MN010 like '%迴轉%' or MN010  like '%評價%' or MN010 like '%手動結案%' or MN010 like '%成本結轉%')
            //            or (left(MN001,1) in ('3','4','5','6'))) and MN020 = 'y' and (MA003 like '%關係%' or MA003 like '%TOPPAN%')";
            //20210315 cond_y 與 cond_N 加入【or (MN001 in ('2218','2219','2166') and MN018 <> 0)】
            cond_N = @"MN020 = 'y' and (MA003 like '%關係%' or MA003 like '%TOPPAN%')
                          and ( MN010 like '%迴轉%' or MN010 like '%評價%' or MN010 like '%手動結案%' or MN010 like '%成本結轉%'
                          or (MN001 in ('2218','2219','2166') and MN018 <> 0))";
            //20210315 sql_Details_Primary 與 sql_Quarter_Primary 更新備註欄位 從科目代號改為 依備註內單別[710T.C71T]判別
            sql_Details_Primary = @"select MN001 as 科目編號, MA003 as 科目名稱, MN003 as 立沖帳目代號
                                    , (case MA012	
                                        when '1' then (select MA124 from COPMA where MA001 = MN003)
                                        when '2' then (select MA085 from PURMA where MA001 = MN003)
                                        End) as 關係人代號
                                    , MN015 as 關係人, MN004 as 傳票日期, MN005 as 傳票單別, MN006 as 傳票單號, MN007 as 序號
                                    , MN011 as 部門代號, ME002 as 部門名稱, MN008 as 借貸別, TB010+TB012 as 摘要
                                    ,(case 
                                        when (TB010+TB012) like '%C71T%' then (Rtrim(MN005) + '-' +Rtrim(MN006))+ ' '+(TB010+TB012)
                                        when (TB010+TB012) like '%710T%' then (Rtrim(MN005) + '-' +Rtrim(MN006))+ ' '+(TB010+TB012)
                                        when (TB010+TB012) like '%股利%' then (Rtrim(MN005) + '-' +Rtrim(MN006))+ ' '+(TB010+TB012)
                                        Else '' End) as 備註
                                    , MN012 as 幣別, MN013 as 匯率
                                    ,(case MN008 when '1' then MN009 else 0 end) as 本幣借方金額
                                    ,(case MN008 when '-1' then MN009 else 0 end) as 本幣貸方金額 
                                    , MN008*MN009 as 本幣金額
                                    ,(case MN008 when '1' then MN014 else 0 end) as 原幣借方金額
                                    ,(case MN008 when '-1' then MN014 else 0 end) as 原幣貸方金額 
                                    , MN008*MN014 as 原幣金額, MN018 as 已沖本幣金額, MN019 as 已沖原幣金額, MN020 as 結案碼
                                  from ACTMN
                                    left join ACTMA on MA001 = MN001
                                    left join ACTTB on TB001 = MN005 and TB002 = MN006 and TB003 = MN007
                                    left join CMSME on ME001 = MN011";

            //cond_Month_A = @"left(MN001,1) in ('1','2') and MN020 = 'N' and (MA003 like '%關係%' or MA003 like '%TOPPAN%')";
            //cond_Month_B = @"left(MN001,1) in ('3','4','5','6') and MN020 = 'N' and (MA003 like '%關係%' or MA003 like '%TOPPAN%')";
            //cond_Quarter = @"MN020 = 'N' and (MA003 like '%關係%' or MA003 like '%TOPPAN%')";

            //20210413 科目單月條件異動，加入科目7開頭，並分為left(MN001,1) in ('1','2','3','7') and MN020 = 'N' and MA012 <> '' and MN004 <= '20210331' 與
            //left(MN001, 1) in ('4', '5', '6') and MN020 = 'N' and MA012<> '' and MN004 like '202103%'，其中3.7 改為累計方式呈現
            cond_Month_A = @"left(MN001,1) in ('1','2','3','7') and MN020 = 'N' and MA012 <> ''";
            cond_Month_B = @"left(MN001,1) in ('4','5','6') and MN020 = 'N' and MA012 <> ''";
            cond_Quarter = @"MN020 = 'N' and  MA012 <> ''";

            //20210315 sql_Details_Primary 與 sql_Quarter_Primary 更新備註欄位 從科目代號改為 依備註內單別[710T.C71T]判別
            sql_Quarter_Primary = @"select MN001 as 科目編號, MA003 as 科目名稱, MN012 as 幣別
                                    , (case MA012	
                                        when '1' then (select MA124 from COPMA where MA001 = MN003)
                                        when '2' then (select MA085 from PURMA where MA001 = MN003)
                                        End) as 關係人代號
                                    , MN015 as 關係人, sum(MN008*MN014) as 原幣金額, sum(MN008*MN009) as 本幣金額
                                    ,(case 
                                        when (TB010+TB012) like '%C71T%' then (Rtrim(MN005) + '-' +Rtrim(MN006))+ ' '+(TB010+TB012)
                                        when (TB010+TB012) like '%710T%' then (Rtrim(MN005) + '-' +Rtrim(MN006))+ ' '+(TB010+TB012)
                                        when (TB010+TB012) like '%股利%' then (Rtrim(MN005) + '-' +Rtrim(MN006))+ ' '+(TB010+TB012)
                                        Else '' End) as 備註
                                     from ACTMN
                                         left join ACTMA on MA001 = MN001
                                         left join ACTTB on TB001 = MN005 and TB002 = MN006 and TB003 = MN007
                                         left join CMSME on ME001 = MN011";

            txt_Preperr.Text = string.Format(@"1.取[結束]抓取月份，例如：2021/01/29，將抓取[2021/01]資訊。
2.日期變更後，先前查詢資料須重新查詢，若無查詢，禁止Excel轉出。
3.請注意!! 建議每月第四個工作天下午後執行。
4.查詢條件：
跨年度僅能每年01月執行。
======== 過帳與成本結轉確認 ===========
先確認傳票(ACTTA)已確認單據(TA010)，全部已過帳(TA011)，扣除915單別；
再確認立沖帳(ACTMN)內摘要有【成本結轉】才能執行指定結案。
{0}
======== 指定結案 N -> y ===========
{1}
======== 取消結案 y -> N ===========
加入修改日期判別 ACTMN.MODI_DATE like 修改年月，無法跨月使用。
{2}", cond_Posting, cond_y, cond_N);

            txterr.Text = string.Format(@"1.取[結束]抓取月份，例如：2021/01/29，將抓取[2021/01]資訊。
2.日期變更後，先前查詢資料須重新查詢，若無查詢，禁止Excel轉出。
3.Excel轉出後包含明細，程式自動開啟該報表。
4.查詢條件：
查詢[立沖帳維護作業]內未結案、關係人代號<>''，
彙總內的備註(TB010+TB012)，採摘要內容內辨別[710T.C71T.股利]關鍵字列出備註
======== 單月 ===========
{0}
{1}
======== 季度 ===========
EXCEL轉出，若月份遇到03.06.09.12，將改為Q1.Q2.Q3.Q4。
{2}", cond_Month_A, cond_Month_B, cond_Quarter);

        }

        private void Btn_Prep_date_s_Click(object sender, EventArgs e)
        {
            str_Prep_date_s = txt_Prep_date_s.Text.Trim();
            this.fm_月曆 = new 月曆(this.txt_Prep_date_s, this.Btn_Prep_date_s, "單據起始日期");
            Prep_DtAndDgvClear();
            chk_LastY3456.Checked = false;
        }

        private void Btn_Prep_date_e_Click(object sender, EventArgs e)
        {
            str_Prep_date_e = txt_Prep_date_e.Text.Trim();
            this.fm_月曆 = new 月曆(this.txt_Prep_date_e, this.Btn_Prep_date_e, "單據結束日期");
            Prep_DtAndDgvClear();
            chk_LastY3456.Checked = false;
        }
        private void btn_Prep_down_Click(object sender, EventArgs e)
        {
            Prep_date_s = DateTime.ParseExact(txt_Prep_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            Prep_date_e = DateTime.ParseExact(txt_Prep_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_Prep_date_s.Text = DateTime.Parse(Prep_date_s.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_Prep_date_e.Text = DateTime.Parse(Prep_date_e.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");

            Prep_DtAndDgvClear();
            chk_LastY3456.Checked = false;
        }

        private void btn_Prep_up_Click(object sender, EventArgs e)
        {
            Prep_date_s = DateTime.ParseExact(txt_Prep_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            Prep_date_e = DateTime.ParseExact(txt_Prep_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_Prep_date_s.Text = DateTime.Parse(Prep_date_s.ToString("yyyy-MM-01")).AddMonths(1).ToString("yyyyMMdd");
            txt_Prep_date_e.Text = DateTime.Parse(Prep_date_e.ToString("yyyy-MM-01")).AddMonths(2).AddDays(-1).ToString("yyyyMMdd");

            Prep_DtAndDgvClear();
            chk_LastY3456.Checked = false;
        }

        private void Btn_date_s_Click(object sender, EventArgs e)
        {
            str_date_s = txt_date_s.Text.Trim();
            this.fm_月曆 = new 月曆(this.txt_date_s, this.Btn_date_s, "單據起始日期");
        }

        private void Btn_date_e_Click(object sender, EventArgs e)
        {
            str_date_e = txt_date_e.Text.Trim();
            this.fm_月曆 = new 月曆(this.txt_date_e, this.Btn_date_e, "單據結束日期");
            str_date_m_e = txt_date_e.Text.Trim().Substring(0, 6);
        }

        private void btn_down_Click(object sender, EventArgs e)
        {
            date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s.Text = DateTime.Parse(date_s.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(date_e.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");

            DtAndDgvClear();
        }

        private void btn_up_Click(object sender, EventArgs e)
        {
            date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s.Text = DateTime.Parse(date_s.ToString("yyyy-MM-01")).AddMonths(1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(date_e.ToString("yyyy-MM-01")).AddMonths(2).AddDays(-1).ToString("yyyyMMdd");

            DtAndDgvClear();
        }
        private void cob_修改者_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.str_修改者ID = this.dt_修改者.Rows[this.cob_修改者.SelectedIndex]["MF001"].ToString().Trim();
            this.str_修改者GP = this.dt_修改者.Rows[this.cob_修改者.SelectedIndex]["MF004"].ToString().Trim();
        }


        private void chk_LastY3456_CheckedChanged(object sender, EventArgs e)
        {
            if (btn_y.Enabled == true || btn_N.Enabled == true) 
            {
                Prep_DtAndDgvClear();
                BtnFalse();
                MessageBox.Show("變更【查詢條件】，請重新【查詢】", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            Prep_date_e = DateTime.ParseExact(txt_Prep_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            str_Prep_date_LastY = DateTime.Parse(Prep_date_e.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");

            //TODO:跨年度查詢條件，需限定每年01月使用
            if (chk_LastY3456.Checked)
            {
                if (str_Prep_date_LastY.Substring(4, 2) == "12")
                {
                    cond_LastY3456 = String.Format(@"or (left(MN001,1) in ('3','4','5','6') and MN020 = 'N' 
                                                    and (MA003 like '%關係%' or MA003 like '%TOPPAN%') and MN004 <= '{0}')", str_Prep_date_LastY);

                    cond_LastN3456 = String.Format(@"or (left(MN001,1) in ('3','4','5','6') and MN020 = 'y' 
                                                    and (MA003 like '%關係%' or MA003 like '%TOPPAN%') and MN004 <= '{0}')", str_Prep_date_LastY);

                }
                else
                {
                    chk_LastY3456.Checked = false;
                    MessageBox.Show("限【查詢條件】為【每年一月】使用", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            else
            {
                cond_LastY3456 = "";
                cond_LastN3456 = "";
            }
        }
        private void btn_PrepSearch_Click(object sender, EventArgs e)
        {
            if (MyClass.DateIntervalCheck(txt_Prep_date_s, txt_Prep_date_e) is false)
            {
                return;
            }

            Prep_DtAndDgvClear();
            BtnFalse();

            str_Prep_date_s = txt_Prep_date_s.Text.Trim();
            str_Prep_date_ym_s = txt_Prep_date_s.Text.Trim().Substring(0, 6);
            str_Prep_date_e = txt_Prep_date_e.Text.Trim();
            str_Prep_date_ym_e = txt_Prep_date_e.Text.Trim().Substring(0, 6);

            //TODO:查詢 傳票是否已過帳
            string sql_str_Posting = String.Format(@"select TA011 as 過帳碼,TA001 as 傳票單別, TA002 as 傳票單號, TA003 as 傳票日期
                        , (case TA006
                            when '1' then '1.一般傳票輸入' when '2' then '2.應計傳票輸入' when '3' then '3.應計回轉'
                            when '4' then '4.常用傳票複製' when '5' then '5.比率分攤' when '6' then '6.迴轉傳票'
                            when '8' then '8.其他轉入' when '7' then '7.紅字沖銷傳票' when 'A' then 'A.票據系統產生'
                            when 'B' then 'B.固定資產產生' when 'C' then 'C.應收系統產生' when 'D' then 'D.應付系統產生'
                            when 'E' then 'E.庫存系統產生' when 'F' then 'F.訂單系統產生' when 'G' then 'G.採購系統產生'
                            when 'H' then 'H.製令系統產生' when 'J' then 'J.專櫃系統產生' when 'K' then 'K.零用金系統產生'
                            when 'L' then 'L.成本計算系統' when 'M' then 'M.幣別轉換' when 'N' then 'N.年結轉 '
                            when 'O' then 'O.遞延收入產生' when 'P' then 'P.迴轉金產生' when 'Q' then 'Q.IFRS開帳傳票'
                            when '9' then '9.月結轉' when 'R' then 'R.IFRS總帳結轉開帳傳票' when 'S' then 'S.IFRS多帳本結轉開帳傳票 '
                            when 'T' then 'T.IFRS開帳調整傳票' when 'U' then 'U.多帳本彙整開帳傳票' when 'V1' then 'V1.總帳結轉開帳傳票 '
                            when 'V2' then 'V2.多帳本結轉開帳傳票' when 'W' then 'W.轉銷'End)as 來源碼
                        , TA010 as 確認碼, TA014 as 確認日, TA015 as 確認者帳號
                        , (select MF002 from ADMMF where MF001= TA015) as 確認者名稱, TA009 as 備註
                     from ACTTA
                    where {0} and TA003 <= '{1}'", cond_Posting, str_Prep_date_e);
            MyCode.Sql_dt(sql_str_Posting, dt_Posting);

            if (dt_Posting.Rows.Count != 0)
            {
                dgv_PostingCost.DataSource = dt_Posting;
                MessageBox.Show("下述傳票【尚未過帳】，請先完成【過帳】後再執行", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                btn_y.Enabled = false;
                btn_N.Enabled = false;
                return;
            }

            //TODO:查詢立沖帳內是否有【成本結轉】
            string sql_str_Cost_N = String.Format(@"select MN020 as 結案碼,ACTMN.MODI_DATE as 修改日期,ACTMN.MODIFIER as 修改者
                                                  ,MN001 as 科目編號, MA003 as 科目名稱, MN003 as 立沖帳目代號, MN015 as 立沖帳目名稱
                                                  , MN004 as 傳票日期, MN010 as 摘要, MN005 as 傳票單別, MN006 as 傳票單號, MN007 as 序號
                                                  , MN011 as 部門代號, ME002 as 部門名稱, MN008 as 借貸別, MN012 as 幣別, MN013 as 匯率
                                                  ,(case MN008 when '1' then MN009 else 0 end) as 本幣借方金額
                                                  ,(case MN008 when '-1' then MN009 else 0 end) as 本幣貸方金額 
                                                  , MN008*MN009 as 本幣借貸餘
                                                  ,(case MN008 when '1' then MN014 else 0 end) as 原幣借方金額
                                                  ,(case MN008 when '-1' then MN014 else 0 end) as 原幣貸方金額 
                                                  , MN008*MN014 as 原幣借貸餘, MN018 as 已沖本幣金額, MN019 as 已沖原幣金額
                                                 from ACTMN
                                                   left join ACTMA on MA001 = MN001
                                                   left join CMSME on ME001 = MN011
                                                where {0} and MN004 like '{1}%'", cond_Cost_N, str_Prep_date_ym_e);
            MyCode.Sql_dt(sql_str_Cost_N, dt_Cost_N);

            //TODO:查詢立沖帳內是否有【成本結轉】
            string sql_str_Cost_y = String.Format(@"select MN020 as 結案碼,ACTMN.MODI_DATE as 修改日期,ACTMN.MODIFIER as 修改者
                                                  ,MN001 as 科目編號, MA003 as 科目名稱, MN003 as 立沖帳目代號, MN015 as 立沖帳目名稱
                                                  , MN004 as 傳票日期, MN010 as 摘要, MN005 as 傳票單別, MN006 as 傳票單號, MN007 as 序號
                                                  , MN011 as 部門代號, ME002 as 部門名稱, MN008 as 借貸別, MN012 as 幣別, MN013 as 匯率
                                                  ,(case MN008 when '1' then MN009 else 0 end) as 本幣借方金額
                                                  ,(case MN008 when '-1' then MN009 else 0 end) as 本幣貸方金額 
                                                  , MN008*MN009 as 本幣借貸餘
                                                  ,(case MN008 when '1' then MN014 else 0 end) as 原幣借方金額
                                                  ,(case MN008 when '-1' then MN014 else 0 end) as 原幣貸方金額 
                                                  , MN008*MN014 as 原幣借貸餘, MN018 as 已沖本幣金額, MN019 as 已沖原幣金額
                                                 from ACTMN
                                                   left join ACTMA on MA001 = MN001
                                                   left join CMSME on ME001 = MN011
                                                where {0} and MN004 like '{1}%'", cond_Cost_y, str_Prep_date_ym_e);
            MyCode.Sql_dt(sql_str_Cost_y, dt_Cost_y);

            //TODO:查詢立沖帳指定條件，需指定結案項目，將查詢結果指定結案
            string sql_str_y = String.Format(@"select MN020 as 結案碼,ACTMN.MODI_DATE as 修改日期,ACTMN.MODIFIER as 修改者
                                                  ,MN001 as 科目編號, MA003 as 科目名稱, MN003 as 立沖帳目代號, MN015 as 立沖帳目名稱
                                                  , MN004 as 傳票日期, MN010 as 摘要, MN005 as 傳票單別, MN006 as 傳票單號, MN007 as 序號
                                                  , MN011 as 部門代號, ME002 as 部門名稱, MN008 as 借貸別, MN012 as 幣別, MN013 as 匯率
                                                  ,(case MN008 when '1' then MN009 else 0 end) as 本幣借方金額
                                                  ,(case MN008 when '-1' then MN009 else 0 end) as 本幣貸方金額 
                                                  , MN008*MN009 as 本幣借貸餘
                                                  ,(case MN008 when '1' then MN014 else 0 end) as 原幣借方金額
                                                  ,(case MN008 when '-1' then MN014 else 0 end) as 原幣貸方金額 
                                                  , MN008*MN014 as 原幣借貸餘, MN018 as 已沖本幣金額, MN019 as 已沖原幣金額
                                                 from ACTMN
                                                   left join ACTMA on MA001 = MN001
                                                   left join CMSME on ME001 = MN011
                                                where ({0} and MN004 <= '{1}') {2}", cond_y, str_Prep_date_e, cond_LastY3456);
            MyCode.Sql_dgv(sql_str_y,dt_y ,dgv_y);

            //TODO:查詢立沖帳指定條件，需指定結案項目，將查詢結果取消結案，加入修改日期判別
            string sql_str_N = String.Format(@"select MN020 as 結案碼,ACTMN.MODI_DATE as 修改日期,ACTMN.MODIFIER as 修改者
                                                  ,MN001 as 科目編號, MA003 as 科目名稱, MN003 as 立沖帳目代號, MN015 as 立沖帳目名稱
                                                  , MN004 as 傳票日期, MN005 as 傳票單別, MN006 as 傳票單號, MN007 as 序號, MN010 as 摘要
                                                  , MN011 as 部門代號, ME002 as 部門名稱, MN008 as 借貸別, MN012 as 幣別, MN013 as 匯率
                                                  ,(case MN008 when '1' then MN009 else 0 end) as 本幣借方金額
                                                  ,(case MN008 when '-1' then MN009 else 0 end) as 本幣貸方金額 
                                                  , MN008*MN009 as 本幣借貸餘
                                                  ,(case MN008 when '1' then MN014 else 0 end) as 原幣借方金額
                                                  ,(case MN008 when '-1' then MN014 else 0 end) as 原幣貸方金額 
                                                  , MN008*MN014 as 原幣借貸餘, MN018 as 已沖本幣金額, MN019 as 已沖原幣金額
                                                 from ACTMN
                                                   left join ACTMA on MA001 = MN001
                                                   left join CMSME on ME001 = MN011
                                                where (({0} {2}) and (ACTMN.MODI_DATE like '{1}%')) ", cond_N, str_修改日期.Substring(0,6), cond_LastN3456);
            MyCode.Sql_dgv(sql_str_N, dt_N, dgv_N);

            if (dt_Cost_N.Rows.Count == 0 && dt_Cost_y.Rows.Count == 0)
            {
                MessageBox.Show("立沖帳尚未查看到【成本結轉】，請確認完成【成本結轉】作業後再執行", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                btn_y.Enabled = false;
                btn_N.Enabled = false;
                return;
            }
            else if (dt_y.Rows.Count == 0 && dt_N.Rows.Count > 0)
            {
                MessageBox.Show("立沖帳本月已有【指定結案(y)】，僅能【取消結案(N)】", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                dgv_PostingCost.DataSource = dt_Cost_y;
                tab_Prep_dgv.SelectedIndex = 2;
                btn_y.Enabled = false;
                btn_N.Enabled = true;
            }
            else if (dt_y.Rows.Count > 0 && dt_N.Rows.Count > 0)
            {
                MessageBox.Show("立沖帳本月已有【指定結案(y)與未結案(N)】", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                dt_Cost_N.Merge(dt_Cost_y);
                dgv_PostingCost.DataSource = dt_Cost_N;
                tab_Prep_dgv.SelectedIndex = 0;
                btn_y.Enabled = true;
                btn_N.Enabled = true;
            }
            else if (dt_y.Rows.Count == 0 && dt_N.Rows.Count == 0)
            {
                MessageBox.Show("立沖帳本月無【未結案(N)】", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                dt_Cost_N.Merge(dt_Cost_y);
                dgv_PostingCost.DataSource = dt_Cost_N;
                tab_Prep_dgv.SelectedIndex = 1;
                btn_y.Enabled = false;
                btn_N.Enabled = false;
            }
            else
            {
                dgv_PostingCost.DataSource = dt_Cost_N;
                tab_Prep_dgv.SelectedIndex = 1;
                btn_y.Enabled = true;
                btn_N.Enabled = false;
                //check_Prep = 1;
            }

        }

        private void btn_y_Click(object sender, EventArgs e)
        {
            if (dgv_y.Rows.Count != 0)
            {
                DialogResult Result = MessageBox.Show("請再次確認資料，將【指定結案(y)】", "確認更新ERP立沖帳狀態", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                if (Result == DialogResult.OK)
                {
                    string sql_str_UpdateToy = String.Format(@"update ACTMN
                            set MN020 ='y',ACTMN.MODIFIER='{2}',ACTMN.MODI_DATE='{3}',ACTMN.FLAG=ACTMN.FLAG+1
                            from ACTMN left join ACTMA on MA001 = MN001
                            where ({0} {1}) ", cond_y, cond_LastY3456, str_修改者ID, str_修改日期);

                    MyCode.sqlExecuteNonQuery(sql_str_UpdateToy, "S2008X64");

                    MessageBox.Show("立沖帳本月已完成【指定結案】", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    txt_Preperr.Text += Environment.NewLine +
                                DateTime.Now.ToString() + Environment.NewLine +
                                ">> 立沖帳本月已完成【指定結案】" + Environment.NewLine +
                                "===========";

                    ////sqlapp log
                    //str_sql_log = String.Format(
                    //          @"insert into develop_app_log VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')"
                    //          , str_建立者ID, str_建立日期, dt_單頭.Rows[x]["TC001"], dt_單頭.Rows[x]["TC002"], "COPTC", "客戶訂單 匯入", "新增客戶訂單單頭", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

                    btn_PrepSearch.PerformClick();
                }
                else if (Result == DialogResult.Cancel)
                {
                    return;
                }

            }
        }


        private void btn_N_Click(object sender, EventArgs e)
        {
            if(dgv_N.Rows.Count != 0) 
            {
                DialogResult Result = MessageBox.Show("請再次確認資料，將【取消指定結案(N)】", "確認更新ERP立沖帳狀態", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                if (Result == DialogResult.OK)
                {
                    string sql_str_UpdateToN = String.Format(@"update ACTMN
                            set MN020 ='N',ACTMN.MODIFIER='{3}',ACTMN.MODI_DATE='{4}',ACTMN.FLAG=ACTMN.FLAG+1
                            from ACTMN left join ACTMA on MA001 = MN001
                            where (({0} {2}) and (ACTMN.MODI_DATE like '{1}%')) ", cond_N, str_修改日期.Substring(0, 6), cond_LastN3456
                            , str_修改者ID, str_修改日期);

                    MyCode.sqlExecuteNonQuery(sql_str_UpdateToN, "S2008X64");

                    MessageBox.Show("立沖帳本月已【取消指定結案】", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    txt_Preperr.Text += Environment.NewLine +
                                DateTime.Now.ToString() + Environment.NewLine +
                                ">> 立沖帳本月已【取消指定結案】" + Environment.NewLine +
                                "===========";
                    btn_PrepSearch.PerformClick();
                }
                else if (Result == DialogResult.Cancel)
                {
                    return;
                }
                   
            }
        }

        private void btn_search_Click(object sender, EventArgs e)
        {
            if (MyClass.DateIntervalCheck(txt_date_s, txt_date_e) is false)
            {
                return;
            }

            DtAndDgvClear();

            str_date_s = txt_date_s.Text.Trim();
            str_date_ym_s = txt_date_s.Text.Trim().Substring(0, 6);
            str_date_e = txt_date_e.Text.Trim();
            str_date_ym_e = txt_date_e.Text.Trim().Substring(0, 6);

            //btn_PrepSearch.PerformClick();

            //TODO:查詢立沖帳內是否有【成本結轉】
            string sql_str_Cost_y = String.Format(@"select MN020 as 結案碼,ACTMN.MODI_DATE as 修改日期,ACTMN.MODIFIER as 修改者
                                                  ,MN001 as 科目編號, MA003 as 科目名稱, MN003 as 立沖帳目代號, MN015 as 立沖帳目名稱
                                                  , MN004 as 傳票日期, MN010 as 摘要, MN005 as 傳票單別, MN006 as 傳票單號, MN007 as 序號
                                                  , MN011 as 部門代號, ME002 as 部門名稱, MN008 as 借貸別, MN012 as 幣別, MN013 as 匯率
                                                  ,(case MN008 when '1' then MN009 else 0 end) as 本幣借方金額
                                                  ,(case MN008 when '-1' then MN009 else 0 end) as 本幣貸方金額 
                                                  , MN008*MN009 as 本幣借貸餘
                                                  ,(case MN008 when '1' then MN014 else 0 end) as 原幣借方金額
                                                  ,(case MN008 when '-1' then MN014 else 0 end) as 原幣貸方金額 
                                                  , MN008*MN014 as 原幣借貸餘, MN018 as 已沖本幣金額, MN019 as 已沖原幣金額
                                                 from ACTMN
                                                   left join ACTMA on MA001 = MN001
                                                   left join CMSME on ME001 = MN011
                                                where {0} and MN004 like '{1}%'", cond_Cost_y, str_date_ym_e);
            MyCode.Sql_dt(sql_str_Cost_y, dt_Cost_y);

            if (dt_Cost_y.Rows.Count == 0) 
            {
                MessageBox.Show("報表查詢月份找不到【成本結轉】，請先執行【指定結案】作業", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            ////銀行存款明細帳細項_評價後及匯入暫存表CT_F22_1_SGLDT_After_Temp
            //string sql_str_Insert_CT_F22_1_SGLDT_After_Temp = String.Format(@"
            //        ", str_date_ym_e);
            //MyCode.sqlExecuteNonQuery(sql_str_Insert_CT_F22_1_SGLDT_After_Temp, "S2008X64");

            //TODO:報表-單月明細
            string sql_str_MonthDetails = String.Format(@"{0}
                            where {1} and MN004 <= '{2}'
                        union all
                            {0}
                            where {3} and MN004 like '{4}%'
                        order by MN001,MN004", sql_Details_Primary, cond_Month_A, str_date_e, cond_Month_B, str_date_ym_s);
            MyCode.Sql_dgv(sql_str_MonthDetails, dt_MonthDetails, dgv_MonthDetails);

            //TODO:報表-累計明細
            string sql_str_QuarterDetails = String.Format(@"{0}
                            where {1} and MN004 <= '{2}'", sql_Details_Primary, cond_Quarter, str_date_e);
            MyCode.Sql_dgv(sql_str_QuarterDetails, dt_QuarterDetails, dgv_QuarterDetails);

            //TODO:報表-單月彙總
            string sql_str_MonthSum = String.Format(@"select 科目編號,科目名稱,幣別,關係人代號,關係人,sum(原幣金額) as 原幣金額,sum(本幣金額) as 本幣金額,備註 
                         from ({0} where {1} and MN004 <= '{2}'
                                group by MN001,MA003,MN015,MN012,ACTMN.MN005,ACTMN.MN006,ACTTB.TB010,ACTTB.TB012,ACTMN.MN003,ACTMA.MA012
                             union all
                                {0}
                                where {3} and MN004 like '{4}%'
                                group by MN001,MA003,MN015,MN012,ACTMN.MN005,ACTMN.MN006,ACTTB.TB010,ACTTB.TB012,ACTMN.MN003,ACTMA.MA012 ) MN_all 
                          where 本幣金額 != 0 and 原幣金額 != 0
                          group by 科目編號,科目名稱,關係人代號,關係人,備註,幣別", sql_Quarter_Primary, cond_Month_A, str_date_e, cond_Month_B, str_date_ym_s);
            MyCode.Sql_dgv(sql_str_MonthSum, dt_MonthSum, dgv_MonthSum);

            //TODO:報表-累計月份彙總
            string sql_str_QuarterSum = String.Format(@"select 科目編號,科目名稱,幣別,關係人代號,關係人,sum(原幣金額) as 原幣金額,sum(本幣金額) as 本幣金額,備註 
                            from ({0} where {1} and MN004 <= '{2}'
                                group by MN001,MA003,MN015,MN012,ACTMN.MN005,ACTMN.MN006,ACTTB.TB010,ACTTB.TB012,ACTMN.MN003,ACTMA.MA012) MN_all
                            where 本幣金額 != 0 and 原幣金額 != 0
                            group by 科目編號,科目名稱,關係人代號,關係人,備註,幣別", sql_Quarter_Primary, cond_Quarter, str_date_e);
            MyCode.Sql_dgv(sql_str_QuarterSum, dt_QuarterSum, dgv_QuarterSum);


            BtnTrue();
        }

        private void dgv_QuarterDetails_DataSourceChanged(object sender, EventArgs e)
        {
            lab_dgv_QuarterDeatails.Text = "共 " + dt_QuarterDetails.Rows.Count.ToString() + " 筆";
        }

        private void dgv_MonthDetails_DataSourceChanged(object sender, EventArgs e)
        {
            lab_dgv_MonthDetails.Text = "共 " + dt_MonthDetails.Rows.Count.ToString() + " 筆";
        }


        private void btn_file_Click(object sender, EventArgs e)
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

        private void btn_fileopen_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process prc = new System.Diagnostics.Process();
            prc.StartInfo.FileName = txt_path.Text.ToString();
            prc.Start();
        }

        private void Prep_DtAndDgvClear()
        {
            dt_Posting.Clear();
            dt_Cost_N.Clear();
            dt_Cost_y.Clear();
            dt_y.Clear();
            dt_N.Clear();
            
            dgv_PostingCost.DataSource = null;
            dgv_y.DataSource = null;
            dgv_N.DataSource = null;
            
            check_Prep = 0;
            BtnFalse();
        }

        private void Prep_BtnFalse()
        {
            btn_y.Enabled = false;
            btn_N.Enabled = false;

        }
        private void Prep_BtnTrue()
        {
            btn_y.Enabled = true;
            btn_N.Enabled = true;

        }

        private void DtAndDgvClear()
        {
            dt_ReportCheckCost.Clear();
            dt_ReportCheckPosting.Clear();
            dt_MonthDetails.Clear();
            dt_QuarterDetails.Clear();
            dt_MonthSum.Clear();
            dt_QuarterSum.Clear();

            dt_Cost_y.Clear();

            dgv_MonthDetails.DataSource = null;
            dgv_QuarterDetails.DataSource = null;
            dgv_MonthSum.DataSource = null;
            dgv_QuarterSum.DataSource = null;

            BtnFalse();
        }

        private void BtnFalse()
        {
            btn_ToExcel.Enabled = false;
        }
        private void BtnTrue()
        {
            btn_ToExcel.Enabled = true;
        }

        private void btn_ToExcel_Click(object sender, EventArgs e)
        {
            BtnFalse();

            using (XLWorkbook wb_RelatedVOU = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel_RelatedVOU))
                {
                    var ws = templateWB.Worksheet("彙總_當月未結案立沖");
                    var ws2 = templateWB.Worksheet("彙總_累計未結案立沖");
                    var ws3 = templateWB.Worksheet("當月未結案立沖");
                    var ws4 = templateWB.Worksheet("累計未結案立沖");

                    ws.CopyTo(wb_RelatedVOU, "彙總_當月未結案立沖");
                    ws2.CopyTo(wb_RelatedVOU, "彙總_累計未結案立沖");
                    ws3.CopyTo(wb_RelatedVOU, "當月未結案立沖");
                    ws4.CopyTo(wb_RelatedVOU, "累計未結案立沖");
                }

                var wsheet_MonthSum = wb_RelatedVOU.Worksheet("彙總_當月未結案立沖");
                var wsheet_QuarterSum = wb_RelatedVOU.Worksheet("彙總_累計未結案立沖");
                var wsheet_MonthDetails = wb_RelatedVOU.Worksheet("當月未結案立沖");
                var wsheet_QuarterDetails = wb_RelatedVOU.Worksheet("累計未結案立沖");

                //=== F22-1_銀行口座一覧表TAST ==========================================
                //wsheet_F22_1_m.Cell(2, 1).Value = "月份區間:" + str_date_ym_s + "~" + str_date_ym_e; //查詢月份區間
                //wsheet_F22_1_m.Cell(3, 1).Value = "製表日期:" + DateTime.Now.ToString("yyyy/MM/dd"); //會計年度
                //wsheet_MonthSum.Cell(3, 1).Value = str_date_ym_e;

                ////== 明細帳(評價前).明細帳(評價後).評價表 =======
                ///ERP_DTInputExcel(wsheet_8aCOPTH, dt_8aCOPTH, str_date_y_e + "01");
                MyCode.ERP_DTInputExcel(wsheet_MonthSum, dt_MonthSum, 5, 1, "RelatedVOU", str_date_ym_e);
                MyCode.ERP_DTInputExcel(wsheet_QuarterSum, dt_QuarterSum, 5, 1, "RelatedVOU", str_date_ym_e);
                MyCode.ERP_DTInputExcel(wsheet_MonthDetails, dt_MonthDetails, 5, 1, str_date_ym_s, str_date_ym_e);
                MyCode.ERP_DTInputExcel(wsheet_QuarterDetails, dt_QuarterDetails, 5, 1, str_date_ym_s.Substring(0,4) + "01", str_date_ym_e);


                save_as_RelatedVOU = txt_path.Text.ToString().Trim() + "\\" + str_date_ym_e + @"_關係人交易_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
                wb_RelatedVOU.SaveAs(save_as_RelatedVOU);

                //打开文件
                if (opencode != 1)
                {
                    System.Diagnostics.Process.Start(save_as_RelatedVOU);
                }
            }
            BtnTrue();
        }

        //private void ERP_DTInputExcel(ClosedXML.Excel.IXLWorksheet wsheet, DataTable dt, int i_col, int j_row, string str_date, string str_SubTotal, string str_Total)
        //{
        //    int i = 0;
        //    int rows_count_dt = dt.Rows.Count;
        //    int col_count_dt = dt.Columns.Count;
        //    string str_SubTotal_Name = "";

        //    wsheet.Cell(2, 2).Value = str_date + "~" + str_date_ym_e; //查詢月份區間
        //    wsheet.Cell(3, 2).Style.NumberFormat.Format = "@";
        //    wsheet.Cell(3, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期

        //    foreach (DataRow row in dt.Rows)
        //    {
        //        int j = 0;

        //        if (str_SubTotal.Length > 0 && str_Total.Length > 0)
        //        {
        //            if (str_SubTotal_Name.ToString() != "" && row[str_SubTotal].ToString() != str_SubTotal_Name.ToString())
        //            {
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
        //                wsheet.Cell(i + i_col, 2).Value = str_SubTotal_Name;
        //                wsheet.Cell(i + i_col, 4).Value = "小計";
        //                wsheet.Range("E" + (i + i_col) + ":K" + (i + i_col)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
        //                wsheet.Cell(i + i_col, 5).FormulaA1 = "=SUMIFS(E:E,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 7).FormulaA1 = "=SUMIFS(G:G,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 8).FormulaA1 = "=SUMIFS(H:H,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 9).FormulaA1 = "=SUMIFS(I:I,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 10).FormulaA1 = "=SUMIFS(J:J,$A:$A,\"" + str_SubTotal_Name + "\")";

        //                i++;
        //            }
        //        }

        //        foreach (DataColumn Column in dt.Columns)
        //        {
        //            switch (Column.ColumnName.ToString())
        //            {
        //                case "銀行帳號":
        //                case "銀行代號":
        //                    wsheet.Cell(i + i_col, j + j_row).Style.NumberFormat.Format = "@";
        //                    break;
        //                case "本幣期初餘額":
        //                case "本幣入帳金額":
        //                case "本幣出帳金額":
        //                case "本幣期末餘額":
        //                case "本幣存款金額":
        //                case "重估本幣金額":
        //                case "匯兌收益":
        //                case "匯兌損失":
        //                    wsheet.Cell(i + i_col, j + j_row).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
        //                    break;
        //                case "原幣進貨金額":
        //                case "原幣期初餘額":
        //                case "原幣入帳金額":
        //                case "原幣出帳金額":
        //                case "原幣期末餘額":
        //                case "原幣存款金額":
        //                case "單位製費成本":
        //                    wsheet.Cell(i + i_col, j + j_row).Style.NumberFormat.Format = "#,##0.00";
        //                    break;
        //                case "匯率":
        //                case "重估匯率":
        //                case "平均匯率":
        //                    wsheet.Cell(i + i_col, j + j_row).Style.NumberFormat.Format = "#,##0.0000";
        //                    break;
        //                default:
        //                    break;
        //            }
        //            wsheet.Cell(i + i_col, j + j_row).Value = row[j];
        //            j++;
        //        }

        //        if (str_SubTotal.Length > 0 && str_Total.Length > 0)
        //        {
        //            str_SubTotal_Name = row[str_SubTotal].ToString().Trim();
        //        }

        //        if ((rows_count_dt - 1) == dt.Rows.IndexOf(row)) //資料列結尾運算
        //        {
        //            if (str_SubTotal.Length == 0 && str_Total.Length > 0)
        //            //if (wsheet.ToString() == "明細帳(評價前)" || wsheet.ToString() == "明細帳(評價後)")
        //            {
        //                i++;
        //                wsheet.Range("A" + (i + i_col) + ":L" + (i + i_col)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":L" + (i + i_col)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":L" + (i + i_col)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
        //                wsheet.Cell(i + i_col, col_count_dt - 2).Value = "小計";
        //                wsheet.Cell(i + i_col, j + j_row - 1).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
        //                wsheet.Cell(i + i_col, j + j_row - 1).FormulaA1 = "=sum(L" + i_col + ":L" + (i + i_col - 1) + ")";
        //                wsheet.SheetView.ZoomScale = 80;

        //            }
        //            //if (wsheet.ToString() == "評價表")
        //            if (str_SubTotal.Length > 0 && str_Total.Length > 0)
        //            {
        //                i++;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
        //                wsheet.Cell(i + i_col, 2).Value = str_SubTotal_Name;
        //                wsheet.Cell(i + i_col, 4).Value = "小計";
        //                wsheet.Range("E" + (i + i_col) + ":K" + (i + i_col)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
        //                wsheet.Cell(i + i_col, 5).FormulaA1 = "=SUMIFS(E:E,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 7).FormulaA1 = "=SUMIFS(G:G,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 8).FormulaA1 = "=SUMIFS(H:H,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 9).FormulaA1 = "=SUMIFS(I:I,$A:$A,\"" + str_SubTotal_Name + "\")";
        //                wsheet.Cell(i + i_col, 10).FormulaA1 = "=SUMIFS(J:J,$A:$A,\"" + str_SubTotal_Name + "\")";

        //                i++;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        //                wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Fill.BackgroundColor = XLColor.Honeydew;
        //                wsheet.Cell(i + i_col, 4).Value = "總計";
        //                wsheet.Range("E" + (i + i_col) + ":K" + (i + i_col)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
        //                wsheet.Cell(i + i_col, 5).FormulaA1 = "=SUMIFS(E:E,$D:$D,\"小計\")";
        //                wsheet.Cell(i + i_col, 7).FormulaA1 = "=SUMIFS(G:G,$D:$D,\"小計\")";
        //                wsheet.Cell(i + i_col, 8).FormulaA1 = "=SUMIFS(H:H,$D:$D,\"小計\")";
        //                wsheet.Cell(i + i_col, 9).FormulaA1 = "=SUMIFS(I:I,$D:$D,\"小計\")";
        //                wsheet.Cell(i + i_col, 10).FormulaA1 = "=SUMIFS(J:J,$D:$D,\"小計\")";
        //                wsheet.Cell(i + i_col, 11).FormulaA1 = "=I" + (i + i_col) + "-J" + (i + i_col);

        //            }
        //        }

        //        i++;
        //    }

        //}


        bool IsToForm1 = false; //紀錄是否要回到Form1
        protected override void OnClosing(CancelEventArgs e) //在視窗關閉時觸發
        {
            ////DialogResult dr = MessageBox.Show("\"是\"回到主畫面 \r\n \"否\"關閉程式", "是否要關閉程式"
            ////    , MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            ////if (dr == DialogResult.Yes)
            ////{
            //IsToForm1 = true;
            ////}
            ////else if (dr == DialogResult.Cancel) 
            ////{

            ////}

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
            Environment.Exit(Environment.ExitCode);
        }

    }
}
