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

namespace TOYOINK_dev
{
    //***********************
    //* 2020/07/08 協理提出新增分頁，公版加入分頁-公司類別分類
    //* 2020/07/10 姿刪提出加入成本調整，加入於明細分類帳，判別科目'510104','510204'摘要內有"成本調整"，歸類於"材料成本"
    //* 2021/02/23 姿刪提出因新增單別2SHT，影響5a及8a自開發報表【銷貨單_勞務收入(佣金)】報表的資料，加入該單別。
    //***********************
    public partial class fm_Package5a8a : Form
    {
        public MyClass MyCode;
        月曆 fm_月曆;

        DataTable dt_8aCOPTH = new DataTable();  //8a品種彙總表
        DataTable dt_5aCOPTH = new DataTable();  //5a明細表
        DataTable dt_COPTH = new DataTable();  //銷貨單
        DataTable dt_COPTJ = new DataTable();  //銷退單
        DataTable dt_INVLA = new DataTable();  //庫存異動單
        DataTable dt_ACTMB = new DataTable();  //損益表
        DataTable dt_ACTML = new DataTable();  //明細分類帳
        DataTable dt_ACRTB = new DataTable();  //銷貨單_勞務收入(佣金)

        DataTable dt_8aCOPTH_m = new DataTable();  //8a品種彙總表
        DataTable dt_5aCOPTH_m = new DataTable();  //5a明細表
        DataTable dt_COPTH_m = new DataTable();  //銷貨單
        DataTable dt_COPTJ_m = new DataTable();  //銷退單
        DataTable dt_INVLA_m = new DataTable();  //庫存異動單
        DataTable dt_ACTMB_m = new DataTable();  //損益表
        DataTable dt_ACTML_m = new DataTable();  //明細分類帳
        DataTable dt_ACRTB_m = new DataTable();  //銷貨單_勞務收入(佣金)

        string createday = DateTime.Now.ToString("yyyy/MM/dd");

        string str_date_s, str_date_m_s;
        string str_date_e, str_date_m_e, str_date_e_ym, str_date_y_e;

        string defaultfilePath = "";
        string cond_8aCOPTH, cond_5aCOPTH, cond_COPTH, cond_COPTJ, cond_INVLA, cond_ACTMB, cond_ACTML, cond_ACRTB;

        string sql_str_8aCOPTH, sql_str_5aCOPTH, sql_str_COPTH, sql_str_COPTJ, sql_str_INVLA, sql_str_ACTMB, sql_str_ACTML, sql_str_ACRTB;
        string sql_str_8aCOPTH_m, sql_str_5aCOPTH_m, sql_str_COPTH_m, sql_str_COPTJ_m, sql_str_INVLA_m, sql_str_ACTMB_m, sql_str_ACTML_m, sql_str_ACRTB_m;

        

        DateTime date_s, date_e;
        string save_as_5aMonth = "", save_as_5aTotal = "", temp_excel_5a, temp_excel_8a, save_as_8aMonth = "", save_as_8aTotal = "";
        int opencode = 0;
        bool err;
        public fm_Package5a8a()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();
            MyCode.strDbCon = "packet size=4096;user id=pwuser;password=sqlmis003;data source=192.168.128.219;persist security info=False;initial catalog=A01A;";
            //MyCode.strDbCon = "packet size=4096;user id=yj.chou;password=yjchou3369;data source=192.168.128.219;persist security info=False;initial catalog=Leader;";
            temp_excel_5a = @"\\192.168.128.219\Company\MIS自開發主檔\會計報表公版\銷貨成本分析月報5a_temp.xlsx";
            temp_excel_8a = @"\\192.168.128.219\Company\MIS自開發主檔\會計報表公版\品種別月報8a_temp.xlsx";


            //MyClass.WriteLog("恢復預設值");
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
        }

        private void btn_up_Click(object sender, EventArgs e)
        {
            date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s.Text = DateTime.Parse(date_s.ToString("yyyy-MM-01")).AddMonths(1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(date_e.ToString("yyyy-MM-01")).AddMonths(2).AddDays(-1).ToString("yyyyMMdd");
        }

        private void DtAndDgvClear()
        {
            //單月
            dt_8aCOPTH_m.Clear();   //8a品種彙總表
            dt_5aCOPTH_m.Clear();   //5a明細表
            dt_COPTH_m.Clear();   //銷貨單
            dt_COPTJ_m.Clear();   //銷退單
            dt_INVLA_m.Clear();   //庫存異動單
            dt_ACTMB_m.Clear();   //損益表
            dt_ACTML_m.Clear();   //明細分類帳
            dt_ACRTB_m.Clear();   //銷貨單_勞務收入(佣金)

            //累計
            dt_8aCOPTH.Clear();   //8a品種彙總表
            dt_5aCOPTH.Clear();   //5a明細表
            dt_COPTH.Clear();   //銷貨單
            dt_COPTJ.Clear();   //銷退單
            dt_INVLA.Clear();   //庫存異動單
            dt_ACTMB.Clear();   //損益表
            dt_ACTML.Clear();   //明細分類帳
            dt_ACRTB.Clear();   //銷貨單_勞務收入(佣金)

            dgv_8aCOPTH.DataSource = null;
            dgv_5aCOPTH.DataSource = null;
            dgv_COPTH.DataSource = null;
            dgv_COPTJ.DataSource = null;

            dgv_INVLA.DataSource = null;
            dgv_ACTMB.DataSource = null;
            dgv_ACTML.DataSource = null;
            dgv_ACRTB.DataSource = null;

            BtnFalse();
        }

        private void BtnFalse()
        {
            btn_5aMonth.Enabled = false;
            btn_5aTotal.Enabled = false;
            btn_5aMT.Enabled = false;

            btn_8aMonth.Enabled = false;
            btn_8aTotal.Enabled = false;
            btn_8aMT.Enabled = false;
        }
        private void BtnTrue()
        {
            btn_5aMonth.Enabled = true;
            btn_5aTotal.Enabled = true;
            btn_5aMT.Enabled = true;

            btn_8aMonth.Enabled = true;
            btn_8aTotal.Enabled = true;
            btn_8aMT.Enabled = true;
        }
        private void txt_date_s_TextChanged(object sender, EventArgs e)
        {
            BtnFalse();

            if (dgv_5aCOPTH.DataSource != null)
            {
                DtAndDgvClear();
            }

        }

        private void txt_date_e_TextChanged(object sender, EventArgs e)
        {
            BtnFalse();

            if (dgv_5aCOPTH.DataSource != null)
            {
                DtAndDgvClear();
            }
        }

        private void fm_Package5a8a_Load(object sender, EventArgs e)
        {
            txt_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
            string filder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            txt_path.Text = filder;

            cond_5aCOPTH = @"COPTG.TG001 IN ('230','2302','232','2301','2303','C230','C231','C234','234T','230T','C23T','2305','235T','2310','2311','2312','2320','2321','2322') AND COPTG.TG023 = 'Y' AND (COPTH.TH007 <> N'43')";
            cond_COPTH = @"TG001 <> '233' and TH007<> '43'";
            cond_COPTJ = @"LA006 IN ('240','241','242','242','243','C240','C241','C242','C244','240T','241T','242T','243T','U240','U241','U242','U24T')";
            cond_INVLA = @"LA006 in ('1105','116','170','173','C110')";
            cond_ACTMB = @"(MB001 like '4%' or MB001 like '51%')";
            cond_ACTML = @"";
            cond_ACRTB = @"TH020 = 'Y' and TH026 = 'Y' and TB004 <>'9' and TH001 in ('C2SH','2SH','2SHT')";

            txterr.Text = string.Format(
                @"1.取[結束]抓取月份，例如：2020/02/29，將抓取[2020/02]資訊。
2.日期變更後，先前查詢資料須重新查詢，若無查詢，禁止Excel轉出。
3.Excel轉出後包含明細，程式自動開啟該報表。
4.查詢條件：
= 5a明細表.8a品種彙總表 5aCOPTH 8aCOPTH=
{0}
========   銷貨單  COPTH ===========
{1}
========   銷退單  COPTJ ===========
{2}
=======  庫存異動單 INVLA ==========
{3}
========   損益表 ACTMB ============
{4}
===== 銷貨單_勞務收入(佣金) ACRTB ====
{5}", cond_5aCOPTH, cond_COPTH, cond_COPTJ, cond_INVLA, cond_ACTMB, cond_ACRTB);
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

        private void btn_search_Click(object sender, EventArgs e)
        {
            if (MyClass.DateIntervalCheck(txt_date_s, txt_date_e) is false)
            {
                return;
            }

            DtAndDgvClear();

            str_date_s = txt_date_s.Text.Trim();
            str_date_m_s = txt_date_s.Text.Trim().Substring(0, 6);
            str_date_e = txt_date_e.Text.Trim();
            str_date_m_e = txt_date_e.Text.Trim().Substring(0, 6);
            str_date_y_e = txt_date_e.Text.Trim().Substring(0, 4);

            //dt_8aCOPTH.Clear();   //8a品種彙總表
            //dt_5aCOPTH.Clear();   //5a明細表
            //dt_COPTH.Clear();   //銷貨單
            //dt_COPTJ.Clear();   //銷退單
            //dt_INVLA.Clear();   //庫存異動單
            //dt_ACTMB.Clear();   //損益表
            //dt_ACTML.Clear();   //明細分類帳
            //dt_ACRTB.Clear();   //銷貨單_勞務收入(佣金)

            if (err == false)
            {
                //單月查詢語法
                SqlCodeSearch(str_date_m_s, str_date_m_e);
                    sql_str_8aCOPTH_m = sql_str_8aCOPTH;
                    sql_str_5aCOPTH_m = sql_str_5aCOPTH;
                    sql_str_COPTH_m = sql_str_COPTH;
                    sql_str_COPTJ_m = sql_str_COPTJ;
                    sql_str_INVLA_m = sql_str_INVLA;
                    sql_str_ACTMB_m = sql_str_ACTMB;
                    sql_str_ACTML_m = sql_str_ACTML;
                    sql_str_ACRTB_m = sql_str_ACRTB;

                //查詢放於dt
                MyCode.Sql_dt(sql_str_8aCOPTH_m, dt_8aCOPTH_m);
                MyCode.Sql_dt(sql_str_5aCOPTH_m, dt_5aCOPTH_m);
                MyCode.Sql_dt(sql_str_COPTH_m, dt_COPTH_m);
                MyCode.Sql_dt(sql_str_COPTJ_m, dt_COPTJ_m);
                MyCode.Sql_dt(sql_str_INVLA_m, dt_INVLA_m);
                MyCode.Sql_dt(sql_str_ACTMB_m, dt_ACTMB_m);
                MyCode.Sql_dt(sql_str_ACTML_m, dt_ACTML_m);
                MyCode.Sql_dt(sql_str_ACRTB_m, dt_ACRTB_m);

                //彙總查詢語法
                SqlCodeSearch(str_date_y_e + "01", str_date_m_e);

                //顯示於dgv
                MyCode.Sql_dgv(sql_str_8aCOPTH, dt_8aCOPTH, dgv_8aCOPTH);
                MyCode.Sql_dgv(sql_str_5aCOPTH, dt_5aCOPTH, dgv_5aCOPTH);
                MyCode.Sql_dgv(sql_str_COPTH, dt_COPTH, dgv_COPTH);
                MyCode.Sql_dgv(sql_str_COPTJ, dt_COPTJ, dgv_COPTJ);
                MyCode.Sql_dgv(sql_str_INVLA, dt_INVLA, dgv_INVLA);
                MyCode.Sql_dgv(sql_str_ACTMB, dt_ACTMB, dgv_ACTMB);
                MyCode.Sql_dgv(sql_str_ACTML, dt_ACTML, dgv_ACTML);
                MyCode.Sql_dgv(sql_str_ACRTB, dt_ACRTB, dgv_ACRTB);
            }
            BtnTrue();
        }


        private void btn_5aMonth_Click(object sender, EventArgs e)
        {
            BtnFalse();

            using (XLWorkbook wb_5aMonth = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel_5a))
                {
                    var ws = templateWB.Worksheet("5a總表");
                    var ws2 = templateWB.Worksheet("5a明細表");
                    var ws3 = templateWB.Worksheet("銷貨單");
                    var ws4 = templateWB.Worksheet("銷退單");
                    var ws5 = templateWB.Worksheet("庫存異動單");
                    var ws6 = templateWB.Worksheet("損益表");
                    var ws7 = templateWB.Worksheet("明細分類帳");
                    var ws8 = templateWB.Worksheet("銷貨單_勞務收入(佣金)");

                    ws.CopyTo(wb_5aMonth, "5a總表");
                    ws2.CopyTo(wb_5aMonth, "5a明細表");
                    ws3.CopyTo(wb_5aMonth, "銷貨單");
                    ws4.CopyTo(wb_5aMonth, "銷退單");
                    ws5.CopyTo(wb_5aMonth, "庫存異動單");
                    ws6.CopyTo(wb_5aMonth, "損益表");
                    ws7.CopyTo(wb_5aMonth, "明細分類帳");
                    ws8.CopyTo(wb_5aMonth, "銷貨單_勞務收入(佣金)");
                }

                var wsheet_5a_m = wb_5aMonth.Worksheet("5a總表");
                var wsheet_5aCOPTH_m = wb_5aMonth.Worksheet("5a明細表");
                var wsheet_COPTH_m = wb_5aMonth.Worksheet("銷貨單");
                var wsheet_COPTJ_m = wb_5aMonth.Worksheet("銷退單");
                var wsheet_INVLA_m = wb_5aMonth.Worksheet("庫存異動單");
                var wsheet_ACTMB_m = wb_5aMonth.Worksheet("損益表");
                var wsheet_ACTML_m = wb_5aMonth.Worksheet("明細分類帳");
                var wsheet_ACRTB_m = wb_5aMonth.Worksheet("銷貨單_勞務收入(佣金)");

                //=== 5a總表 ==========================================
                wsheet_5a_m.Cell(2, 1).Value = "月份區間:" + str_date_m_s + "~" + str_date_m_e; //查詢月份區間
                wsheet_5a_m.Cell(3, 1).Value = "製表日期:" + DateTime.Now.ToString("yyyy/MM/dd"); //會計年度

                //== 5a明細表 銷貨單 銷退單 庫存異動單 損益表 明細分類帳 銷貨單_勞務收入(佣金) =======
                ERP_DTInputExcel(wsheet_5aCOPTH_m, dt_5aCOPTH_m, str_date_m_s);
                ERP_DTInputExcel(wsheet_COPTH_m, dt_COPTH_m, str_date_m_s);
                ERP_DTInputExcel(wsheet_COPTJ_m, dt_COPTJ_m, str_date_m_s);
                ERP_DTInputExcel(wsheet_INVLA_m, dt_INVLA_m, str_date_m_s);
                ERP_DTInputExcel(wsheet_ACTMB_m, dt_ACTMB_m, str_date_m_s);
                ERP_DTInputExcel(wsheet_ACTML_m, dt_ACTML_m, str_date_m_s);
                ERP_DTInputExcel(wsheet_ACRTB_m, dt_ACRTB_m, str_date_m_s);

                save_as_5aMonth = txt_path.Text.ToString().Trim() + "\\" + str_date_m_e + @"銷貨成本分析月報5a_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
                wb_5aMonth.SaveAs(save_as_5aMonth);

                //打开文件
                if (opencode != 1)
                {
                    System.Diagnostics.Process.Start(save_as_5aMonth);
                }
            }
            BtnTrue();
        }
        private void btn_5aTotal_Click(object sender, EventArgs e)
        {
            BtnFalse();

            using (XLWorkbook wb_5aTotal = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel_5a))
                {
                    var ws = templateWB.Worksheet("5a總表");
                    var ws2 = templateWB.Worksheet("5a明細表");
                    var ws3 = templateWB.Worksheet("銷貨單");
                    var ws4 = templateWB.Worksheet("銷退單");
                    var ws5 = templateWB.Worksheet("庫存異動單");
                    var ws6 = templateWB.Worksheet("損益表");
                    var ws7 = templateWB.Worksheet("明細分類帳");
                    var ws8 = templateWB.Worksheet("銷貨單_勞務收入(佣金)");

                    ws.CopyTo(wb_5aTotal, "5a總表");
                    ws2.CopyTo(wb_5aTotal, "5a明細表");
                    ws3.CopyTo(wb_5aTotal, "銷貨單");
                    ws4.CopyTo(wb_5aTotal, "銷退單");
                    ws5.CopyTo(wb_5aTotal, "庫存異動單");
                    ws6.CopyTo(wb_5aTotal, "損益表");
                    ws7.CopyTo(wb_5aTotal, "明細分類帳");
                    ws8.CopyTo(wb_5aTotal, "銷貨單_勞務收入(佣金)");
                }

                var wsheet_5a = wb_5aTotal.Worksheet("5a總表");
                var wsheet_5aCOPTH = wb_5aTotal.Worksheet("5a明細表");
                var wsheet_COPTH = wb_5aTotal.Worksheet("銷貨單");
                var wsheet_COPTJ = wb_5aTotal.Worksheet("銷退單");
                var wsheet_INVLA = wb_5aTotal.Worksheet("庫存異動單");
                var wsheet_ACTMB = wb_5aTotal.Worksheet("損益表");
                var wsheet_ACTML = wb_5aTotal.Worksheet("明細分類帳");
                var wsheet_ACRTB = wb_5aTotal.Worksheet("銷貨單_勞務收入(佣金)");

                //=== 5a總表 ==========================================
                wsheet_5a.Cell(2, 1).Value = "月份區間:" + str_date_y_e + "01" + "~" + str_date_m_e; //查詢月份區間
                wsheet_5a.Cell(3, 1).Value = "製表日期:" + DateTime.Now.ToString("yyyy/MM/dd"); //會計年度

                ////== 5a明細表 銷貨單 銷退單 庫存異動單 損益表 明細分類帳 銷貨單_勞務收入(佣金) =======
                ERP_DTInputExcel(wsheet_5aCOPTH, dt_5aCOPTH, str_date_y_e + "01");
                ERP_DTInputExcel(wsheet_COPTH, dt_COPTH, str_date_y_e + "01");
                ERP_DTInputExcel(wsheet_COPTJ, dt_COPTJ, str_date_y_e + "01");
                ERP_DTInputExcel(wsheet_INVLA, dt_INVLA, str_date_y_e + "01");
                ERP_DTInputExcel(wsheet_ACTMB, dt_ACTMB, str_date_y_e + "01");
                ERP_DTInputExcel(wsheet_ACTML, dt_ACTML, str_date_y_e + "01");
                ERP_DTInputExcel(wsheet_ACRTB, dt_ACRTB, str_date_y_e + "01");

                save_as_5aTotal = txt_path.Text.ToString().Trim() + "\\" + str_date_y_e + "01-" + str_date_m_e.Substring(4, 2) + @"銷貨成本分析月報5a-彙總表_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
                wb_5aTotal.SaveAs(save_as_5aTotal);

                //打开文件
                if (opencode != 1)
                {
                    System.Diagnostics.Process.Start(save_as_5aTotal);
                }
            }
            BtnTrue();
        }

        private void btn_5aMT_Click(object sender, EventArgs e)
        {
            opencode = 1;

            BtnFalse();

            btn_5aMonth_Click(null, new EventArgs());
            btn_5aTotal_Click(null, new EventArgs());
            System.Diagnostics.Process.Start(save_as_5aMonth);
            System.Diagnostics.Process.Start(save_as_5aTotal);

            opencode = 0;
            //btn_5aMonth.Enabled = true;
            //btn_5aTotal.Enabled = true;
            //btn_5aMT.Enabled = true;
        }

        private void btn_8aMonth_Click(object sender, EventArgs e)
        {
            BtnFalse();

            using (XLWorkbook wb_8aMonth = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel_8a))
                {
                    var ws = templateWB.Worksheet("8a總表");
                    var ws2 = templateWB.Worksheet("8a明細表");
                    var ws3 = templateWB.Worksheet("品種彙總表");
                    var ws4 = templateWB.Worksheet("銷貨單");
                    var ws5 = templateWB.Worksheet("銷退單");
                    var ws6 = templateWB.Worksheet("庫存異動單");
                    var ws7 = templateWB.Worksheet("損益表");
                    var ws8 = templateWB.Worksheet("明細分類帳");
                    var ws9 = templateWB.Worksheet("銷貨單_勞務收入(佣金)");
                    var ws10 = templateWB.Worksheet("公司類別分類");
                    

                    ws.CopyTo(wb_8aMonth, "8a總表");
                    ws2.CopyTo(wb_8aMonth, "8a明細表");
                    ws10.CopyTo(wb_8aMonth, "公司類別分類"); //協理用
                    ws3.CopyTo(wb_8aMonth, "品種彙總表");
                    ws4.CopyTo(wb_8aMonth, "銷貨單");
                    ws5.CopyTo(wb_8aMonth, "銷退單");
                    ws6.CopyTo(wb_8aMonth, "庫存異動單");
                    ws7.CopyTo(wb_8aMonth, "損益表");
                    ws8.CopyTo(wb_8aMonth, "明細分類帳");
                    ws9.CopyTo(wb_8aMonth, "銷貨單_勞務收入(佣金)");
                }

                var wsheet_8a_m = wb_8aMonth.Worksheet("8a總表");
                var wsheet_8a_2_m = wb_8aMonth.Worksheet("8a明細表");
                var wsheet_8a_company = wb_8aMonth.Worksheet("公司類別分類");
                var wsheet_8aCOPTH_m = wb_8aMonth.Worksheet("品種彙總表");
                var wsheet_COPTH_m = wb_8aMonth.Worksheet("銷貨單");
                var wsheet_COPTJ_m = wb_8aMonth.Worksheet("銷退單");
                var wsheet_INVLA_m = wb_8aMonth.Worksheet("庫存異動單");
                var wsheet_ACTMB_m = wb_8aMonth.Worksheet("損益表");
                var wsheet_ACTML_m = wb_8aMonth.Worksheet("明細分類帳");
                var wsheet_ACRTB_m = wb_8aMonth.Worksheet("銷貨單_勞務收入(佣金)");

                //=== 8a總表 ==========================================
                wsheet_8a_m.Cell(2, 1).Value = "月份區間:" + str_date_m_s + "~" + str_date_m_e; //查詢月份區間
                wsheet_8a_m.Cell(3, 1).Value = "製表日期:" + DateTime.Now.ToString("yyyy/MM/dd"); //會計年度
                wsheet_8a_m.SheetView.ZoomScale = 70;

                //=== 8a明細表 ==========================================
                wsheet_8a_2_m.Cell(2, 1).Value = "月份區間:" + str_date_m_s + "~" + str_date_m_e; //查詢月份區間
                wsheet_8a_2_m.Cell(3, 1).Value = "製表日期:" + DateTime.Now.ToString("yyyy/MM/dd"); //會計年度
                wsheet_8a_2_m.SheetView.ZoomScale = 80;

                //=== 公司類別分類 ==========================================
                wsheet_8a_company.Cell(2, 1).Value = "月份區間:" + str_date_m_s + "~" + str_date_m_e; //查詢月份區間
                wsheet_8a_company.Cell(3, 1).Value = "製表日期:" + DateTime.Now.ToString("yyyy/MM/dd"); //會計年度
                wsheet_8a_company.SheetView.ZoomScale = 80;

                //== 品種彙總表 銷貨單 銷退單 庫存異動單 損益表 明細分類帳 銷貨單_勞務收入(佣金) =======
                ERP_DTInputExcel(wsheet_8aCOPTH_m, dt_8aCOPTH_m, str_date_m_s);
                ERP_DTInputExcel(wsheet_COPTH_m, dt_COPTH_m, str_date_m_s);
                ERP_DTInputExcel(wsheet_COPTJ_m, dt_COPTJ_m, str_date_m_s);
                ERP_DTInputExcel(wsheet_INVLA_m, dt_INVLA_m, str_date_m_s);
                ERP_DTInputExcel(wsheet_ACTMB_m, dt_ACTMB_m, str_date_m_s);
                ERP_DTInputExcel(wsheet_ACTML_m, dt_ACTML_m, str_date_m_s);
                ERP_DTInputExcel(wsheet_ACRTB_m, dt_ACRTB_m, str_date_m_s);

                save_as_8aMonth = txt_path.Text.ToString().Trim() + "\\" + str_date_m_e + @"品種別月報8a_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
                wb_8aMonth.SaveAs(save_as_8aMonth);

                //打开文件
                if (opencode != 1)
                {
                    System.Diagnostics.Process.Start(save_as_8aMonth);
                }
            }
            BtnTrue();
        }
        private void btn_8aTotal_Click(object sender, EventArgs e)
        {
            BtnFalse();

            using (XLWorkbook wb_8aTotal = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel_8a))
                {
                    var ws = templateWB.Worksheet("8a總表");
                    var ws2 = templateWB.Worksheet("8a明細表");
                    var ws3 = templateWB.Worksheet("品種彙總表"); 
                    var ws4 = templateWB.Worksheet("銷貨單");
                    var ws5 = templateWB.Worksheet("銷退單");
                    var ws6 = templateWB.Worksheet("庫存異動單");
                    var ws7 = templateWB.Worksheet("損益表");
                    var ws8 = templateWB.Worksheet("明細分類帳");
                    var ws9 = templateWB.Worksheet("銷貨單_勞務收入(佣金)");
                    var ws10 = templateWB.Worksheet("公司類別分類");

                    ws.CopyTo(wb_8aTotal, "8a總表");
                    ws2.CopyTo(wb_8aTotal, "8a明細表");
                    ws10.CopyTo(wb_8aTotal, "公司類別分類"); //協理用
                    ws3.CopyTo(wb_8aTotal, "品種彙總表");
                    ws4.CopyTo(wb_8aTotal, "銷貨單");
                    ws5.CopyTo(wb_8aTotal, "銷退單");
                    ws6.CopyTo(wb_8aTotal, "庫存異動單");
                    ws6.CopyTo(wb_8aTotal, "損益表");
                    ws8.CopyTo(wb_8aTotal, "明細分類帳");
                    ws9.CopyTo(wb_8aTotal, "銷貨單_勞務收入(佣金)");
                    
                }

                var wsheet_8a = wb_8aTotal.Worksheet("8a總表");
                var wsheet_8a_2 = wb_8aTotal.Worksheet("8a明細表");
                var wsheet_8a_company = wb_8aTotal.Worksheet("公司類別分類");
                var wsheet_8aCOPTH = wb_8aTotal.Worksheet("品種彙總表");
                var wsheet_COPTH = wb_8aTotal.Worksheet("銷貨單");
                var wsheet_COPTJ = wb_8aTotal.Worksheet("銷退單");
                var wsheet_INVLA = wb_8aTotal.Worksheet("庫存異動單");
                var wsheet_ACTMB = wb_8aTotal.Worksheet("損益表");
                var wsheet_ACTML = wb_8aTotal.Worksheet("明細分類帳");
                var wsheet_ACRTB = wb_8aTotal.Worksheet("銷貨單_勞務收入(佣金)");

                //=== 8a總表 ==========================================
                wsheet_8a.Cell(2, 1).Value = "月份區間:" + str_date_y_e + "01" + "~" + str_date_m_e; //查詢月份區間
                wsheet_8a.Cell(3, 1).Value = "製表日期:" + DateTime.Now.ToString("yyyy/MM/dd"); //會計年度
                wsheet_8a.SheetView.ZoomScale = 70;

                //=== 8a明細表 ==========================================
                wsheet_8a_2.Cell(2, 1).Value = "月份區間:" + str_date_y_e + "01" + "~" + str_date_m_e; //查詢月份區間
                wsheet_8a_2.Cell(3, 1).Value = "製表日期:" + DateTime.Now.ToString("yyyy/MM/dd"); //會計年度
                wsheet_8a_2.SheetView.ZoomScale = 80;

                //=== 公司類別分類 ==========================================
                wsheet_8a_company.Cell(2, 1).Value = "月份區間:" + str_date_y_e + "01" + "~" + str_date_m_e; //查詢月份區間
                wsheet_8a_company.Cell(3, 1).Value = "製表日期:" + DateTime.Now.ToString("yyyy/MM/dd"); //會計年度
                wsheet_8a_company.SheetView.ZoomScale = 80;

                ////== 品種彙總表 銷貨單 銷退單 庫存異動單 損益表 明細分類帳 銷貨單_勞務收入(佣金) =======
                ERP_DTInputExcel(wsheet_8aCOPTH, dt_8aCOPTH, str_date_y_e + "01");
                ERP_DTInputExcel(wsheet_COPTH, dt_COPTH, str_date_y_e + "01");
                ERP_DTInputExcel(wsheet_COPTJ, dt_COPTJ, str_date_y_e + "01");
                ERP_DTInputExcel(wsheet_INVLA, dt_INVLA, str_date_y_e + "01");
                ERP_DTInputExcel(wsheet_ACTMB, dt_ACTMB, str_date_y_e + "01");
                ERP_DTInputExcel(wsheet_ACTML, dt_ACTML, str_date_y_e + "01");
                ERP_DTInputExcel(wsheet_ACRTB, dt_ACRTB, str_date_y_e + "01");

                save_as_8aTotal = txt_path.Text.ToString().Trim() + "\\" + str_date_y_e + "01-" + str_date_m_e.Substring(4, 2) + @"品種別月報8a-彙總表_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
                wb_8aTotal.SaveAs(save_as_8aTotal);

                //打开文件
                if (opencode != 1)
                {
                    System.Diagnostics.Process.Start(save_as_8aTotal);
                }
            }
            BtnTrue();
        }

        private void btn_8aMT_Click(object sender, EventArgs e)
        {
            opencode = 1;

            BtnFalse();

            btn_8aMonth_Click(null, new EventArgs());
            btn_8aTotal_Click(null, new EventArgs());
            System.Diagnostics.Process.Start(save_as_8aMonth);
            System.Diagnostics.Process.Start(save_as_8aTotal);

            opencode = 0;
            //btn_5aMonth.Enabled = true;
            //btn_5aTotal.Enabled = true;
            //btn_5aMT.Enabled = true;
        }

        private void ERP_DTInputExcel(ClosedXML.Excel.IXLWorksheet wsheet, DataTable dt, string str_date)
        {
            int i = 0;

            wsheet.Cell(2, 2).Value = str_date + "~" + str_date_m_e; //查詢月份區間
            wsheet.Cell(3, 2).Style.NumberFormat.Format = "@";
            wsheet.Cell(3, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期
            foreach (DataRow row in dt.Rows)
            {
                int j = 0;
                foreach (DataColumn Column in dt.Columns)
                {
                    switch (Column.ColumnName.ToString())
                    {
                        case "銷貨年月":
                        case "銷退年月":
                        case "單別":
                        case "會計別":
                        case "產品別":
                        case "商品":
                        case "存貨會計科目":
                        case "年月":
                        case "科目編號":
                        case "科目層級1":
                        case "科目層級2":
                        case "科目層級3":
                        case "會計年度":
                        case "期別":
                        case "年度":
                        case "傳票年月":
                        case "傳票編號":
                        case "單據日期":
                        case "單據年月":
                        case "銷貨單別":
                        case "銷貨單號":
                        case "結帳單別":
                        case "結帳單號":
                        case "結帳序號":
                        case "來源":
                            wsheet.Cell(i + 5, j + 1).Style.NumberFormat.Format = "@";
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
                        case "未稅金額":
                        case "稅額":
                        case "總金額":
                        case "成本":
                        case "材料成本":
                        case "人工成本":
                        case "製費成本":
                        case "利潤":
                        case "罐數":
                            wsheet.Cell(i + 5, j + 1).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                            break;
                        case "銷貨數量":
                        case "銷貨數":
                        case "銷退數":
                            wsheet.Cell(i + 5, j + 1).Style.NumberFormat.Format = "#,##0.000";
                            break;
                        case "平均單價":
                        case "利潤比率":
                        case "平均單位成本":
                        case "單位材料成本":
                        case "單位人工成本":
                        case "單位製費成本":
                            wsheet.Cell(i + 5, j + 1).Style.NumberFormat.Format = "#,##0.00";
                            break;
                        default:
                            break;
                    }
                    wsheet.Cell(i + 5, j + 1).Value = row[j];
                    j++;
                }
                i++;
            }
        }

        private void SqlCodeSearch(string Date_S, string Date_E)
        {
            //8a品種彙總表
            sql_str_8aCOPTH = String.Format(
                    @" SELECT INVMB.MB006 as 產品別
                         , INVMA.MA003 as 商品
                         ,SUM(COPTH.TH008) as 銷貨數量
                         ,SUM(COPTH.TH037) as 未稅金額
                         ,SUM(INVLA.LA017+INVLA.LA018+INVLA.LA019) as 成本
                         ,SUM(INVLA.LA017) as 材料成本
                         ,SUM(INVLA.LA018) as 人工成本
                         ,SUM(INVLA.LA019) as 製費成本
                         ,left(COPTG.TG003,6) as 銷售月份
                        FROM COPTH as COPTH  
                        Inner JOIN COPTG as COPTG On COPTG.TG001=COPTH.TH001 and COPTG.TG002=COPTH.TH002 
                        Inner JOIN INVMB as INVMB On COPTH.TH004=INVMB.MB001 
                        Inner JOIN INVLA as INVLA On COPTH.TH004=INVLA.LA001 and COPTH.TH001=INVLA.LA006 and COPTH.TH002=INVLA.LA007 and COPTH.TH003=INVLA.LA008 
                        left JOIN INVMA as INVMA On INVMA.MA002=INVMB.MB007
                        WHERE ({0}) and left(COPTG.TG003,6) between '{1}' and '{2}'
                        group by MB006,MA003,left(COPTG.TG003,6)
                        order by MB006", cond_5aCOPTH, Date_S, Date_E);
            //5a明細表
            sql_str_5aCOPTH = String.Format(
                @"SELECT INVMB.MB006 as 產品別,
                    INVMA.MA003 as 商品,
                    COPMA.MA002 as 客戶簡稱,
                    COPTH.TH004 as 品號,
                    SUM(COPTH.TH008) as 銷貨數量,
                    INVMB.MB004 as 單位,
                    SUM(COPTH.TH039) as 罐數,
                    SUM(COPTH.TH037)/SUM(COPTH.TH008) as 平均單價,
                    SUM(COPTH.TH037) as 未稅金額,
                    SUM(COPTH.TH038) as 稅額,
                    SUM(COPTH.TH037)+SUM(COPTH.TH038) as 總金額,
                    SUM(INVLA.LA013)/SUM(COPTH.TH008) as 平均單位成本,
                    SUM(INVLA.LA013) as 成本,
                    SUM(INVLA.LA017) as 材料成本,
                    SUM(INVLA.LA018) as 人工成本,
                    SUM(INVLA.LA019) as 製費成本,
                    SUM(COPTH.TH037)-SUM(INVLA.LA013) as 利潤,
                    Case when SUM(COPTH.TH037)<>0 then (SUM(COPTH.TH037)-SUM(INVLA.LA013))/SUM(COPTH.TH037)*100 
                    else 0 end as 利潤比率,
                    SUM(INVLA.LA017)/SUM(COPTH.TH008) as 單位材料成本,
                    SUM(INVLA.LA018)/SUM(COPTH.TH008) as 單位人工成本,
                    SUM(INVLA.LA019)/SUM(COPTH.TH008) as 單位製費成本
                     FROM COPTH as COPTH  
                     Inner JOIN COPTG as COPTG On COPTH.TH001=COPTG.TG001 and COPTH.TH002=COPTG.TG002 
                     Inner JOIN INVLA as INVLA On COPTH.TH001=INVLA.LA006 and COPTH.TH002=INVLA.LA007 and COPTH.TH003=INVLA.LA008 and COPTH.TH004=INVLA.LA001 
                     Inner JOIN INVMB as INVMB On COPTH.TH004=INVMB.MB001 Left JOIN COPMA as COPMA On COPTG.TG004=COPMA.MA001 
                    left join INVMA as INVMA on INVMA.MA002=INVMB.MB007
                    WHERE ({0}) and left(COPTG.TG003,6) between '{1}' and '{2}'
                    GROUP BY INVMB.MB006,INVMA.MA003,INVMB.MB004,COPTH.TH004,COPTG.TG004,COPMA.MA002
                    ORDER BY INVMB.MB006 asc,COPTH.TH004 asc", cond_5aCOPTH, Date_S, Date_E);

            //dt_COPTH.Clear();   //銷貨單
            sql_str_COPTH = String.Format(
                @"SELECT (case when TG006='WCSOT' then 'WCSOT' else COPMA.MA002 end) as 客戶別
                       ,SUBSTRING(TG003,1,6) as 銷貨年月
                       ,TH004 as 品目
                       ,sum(TH008) as 銷貨數
                       ,sum(TH037) as 本幣未稅金額
                       ,sum(LA013) as 總原價
                       ,sum(LA017) as 材料
                       ,sum(LA018) as 人工
                       ,sum(LA019) as 製費
                       ,TG001 as 單別
                       ,MQ002 as 單據名稱
                       ,MB005 as 會計別
                       ,MB006 as 產品別
                       ,INVMA.MA003 as 商品
                     FROM COPTG 
                     INNER JOIN COPTH ON TG001 = TH001 AND TG002 = TH002 
                     INNER JOIN INVLA ON TH001=LA006 and TH002=LA007 and TH003=LA008 and TH004=LA001 
                     left join INVMB ON MB001 = TH004
                     left join INVMA ON MB007=INVMA.MA002 and MB001 = TH004
                     left join CMSMQ on MQ001 = TG001
                     left join COPMA on COPMA.MA001 = TG004
                     WHERE {0} and SUBSTRING(TG003,1,6) between '{1}' and '{2}'
                    group by (case when TG006='WCSOT' then 'WCSOT' else COPMA.MA002 end),
                    SUBSTRING(TG003,1,6),TH004,TG001,MQ002,MB005,MB006,INVMA.MA003", cond_COPTH, Date_S, Date_E);

            //dt_COPTJ.Clear();   //銷退單
            sql_str_COPTJ = String.Format(
                @"SELECT LA010 as 客戶別
                      ,SUBSTRING(LA004,1,6) as 銷退年月
                      ,LA001 as 品目
                      ,sum(LA011) as 銷退數
                      ,sum(TJ033) as 本幣未稅金額
                      ,sum(LA013) as 總原價
                      ,sum(LA017) as 材料 
                      ,sum(LA018) as 人工
                      ,sum(LA019) as 製費
                      ,LA006 as 單別
                      ,MQ002 as 單據名稱
                      ,MB005 as 會計別
                      ,MB006 as 產品別
                      ,MA003 as 商品
                     FROM INVLA 
                     left join INVMB on MB001 = LA001
                     left join COPTJ ON TJ001 = LA006 AND TJ002 = LA007 AND TJ003 = LA008
                     left join INVMA ON INVMB.MB007=MA002 and INVMB.MB001 = LA001
                     left join CMSMQ on MQ001 = LA006
                     WHERE {0} AND SUBSTRING(LA004,1,6) between '{1}' AND '{2}' 
                     group by LA010,SUBSTRING(LA004,1,6),LA001,LA006,MQ002,MB005,MB006,MA003", cond_COPTJ, Date_S, Date_E);

            //dt_INVLA.Clear();   //庫存異動單
            sql_str_INVLA = String.Format(
                @"select 	LA001	品號
                        ,	MB002	品名
                        ,	MB005	會計別
                        ,	MA003	會計別名稱
                        ,	(case when LA006 = '170' and MA004 in ('1211','1218','1220','1226','1227') then '52' when LA006 = '170' and MA004 in ('1233','1234','1235','1236') then '26' else MB006 end) as 產品別
                        ,	MA004	存貨會計科目	
                        ,	SUBSTRING(LA004,1,6) 年月
                        ,	LA006	單別
                        ,	MQ002	單據名稱
                        ,	sum(LA013)	金額
                        ,	sum(LA017)	金額_材料
                        ,	sum(LA018)	金額_人工
                        ,	sum(LA019)	金額_製費
                        ,	sum(LA020)	金額_加工
                        , INVMA.MA003 商品
                        from INVLA
                        left join INVMB on INVMB.MB001=INVLA.LA001
                        left join INVMA as INVMA on INVMA.MA002=INVMB.MB005 and MA001='1'
                        left JOIN CMSMQ on CMSMQ.MQ001  = INVLA.LA006 
                    where {0} and SUBSTRING(LA004,1,6) between '{1}' and '{2}'
                    group by LA001,MB002,MB005,MA003,MB006,MA004,LA004,LA006,MQ002
                    order by LA006", cond_INVLA, Date_S, Date_E);

            ////dt_ACTMB.Clear();   //損益表
            sql_str_ACTMB = String.Format(
                @"SELECT	MB001	as	科目編號
                        ,(select MA003 from [A01A].[dbo].ACTMA where MA001 = ACTMB.MB001) as 科目名稱
                        ,Left(MB001,1) as 科目層級1
                        ,Left(MB001,2) as 科目層級2
                        ,Left(MB001,3) as 科目層級3
                        ,	MB002	as	會計年度
                        ,	MB003	as	期別
                        ,(MB002 + MB003) as 年度
                        ,(case MA007 when '1' then sum(MB004-MB005) when '-1' then sum(MB005-MB004) else 0 end ) 貸借金額
                          FROM ACTMB
                    left JOIN ACTMA on ACTMA.MA001 = ACTMB.MB001
                    where {0} and (MB002 + MB003) between '{1}' and '{2}'
                    group by MB001,MB002,MB003,MA007", cond_ACTMB, Date_S, Date_E);

            ////dt_ACTML.Clear();   //明細分類帳
            sql_str_ACTML = String.Format(
                @"select * from (
                            select ML006 as 科目編號 
                            ,(select MA003 from ACTMA where MA001 = ACTML.ML001) as 科目名稱
                            ,SUBSTRING(ML002,1,6) as 傳票年月 ,ML003+'-'+ML004+' -'+ML005 as 傳票編號
                            ,ML009 as 摘要 ,TB012 as 備註
                            ,(case ML007 when '1' then ML008 else 0 end) as 本幣借方金額
                            ,(case ML007 when '-1' then ML008 else 0 end)  as 本幣貸方金額 
                            ,(case MA007 
	                            when '1' then ((case ML007 when '1' then ML008 else 0 end)-(case ML007 when '-1' then ML008 else 0 end)) 
	                            when '-1' then ((case ML007 when '-1' then ML008 else 0 end)-(case ML007 when '1' then ML008 else 0 end)) else 0 end ) as 貸借金額
                            ,'閒置' as 項目, SUBSTRING(ML009,12,4) as 成本,'52' as 產品別
                            from ACTML
	                            left JOIN ACTMA on ACTMA.MA001 = ACTML.ML006
                                left JOIN ACTTB on ACTTB.TB001 = ACTML.ML003 and ACTTB.TB002 = ACTML.ML004 and ACTTB.TB003 = ACTML.ML005
                            where ML009 like '%閒置%' and ML006 = '510502'
                         UNION ALL
                            select ML006 as 科目編號 
                            ,(select MA003 from ACTMA where MA001 = ACTML.ML001) as 科目名稱
                            ,SUBSTRING(ML002,1,6) as 傳票年月,ML003+'-'+ML004+' -'+ML005 as 傳票編號
                            ,ML009 as 摘要 ,TB012 as 備註
                            ,(case ML007 when '1' then ML008 else 0 end) as 本幣借方金額
                            ,(case ML007 when '-1' then ML008 else 0 end)  as 本幣貸方金額 
                            ,(case MA007 
	                            when '1' then ((case ML007 when '1' then ML008 else 0 end)-(case ML007 when '-1' then ML008 else 0 end)) 
	                            when '-1' then ((case ML007 when '-1' then ML008 else 0 end)-(case ML007 when '1' then ML008 else 0 end)) else 0 end ) as 貸借金額
                            ,'稅額調整' as 項目,'未稅金額' as 成本,SUBSTRING(ML009,CHARINDEX('#',ML009)+1,2) as 產品別
                            from ACTML
	                            left JOIN ACTMA on ACTMA.MA001 = ACTML.ML006
                                left JOIN ACTTB on ACTTB.TB001 = ACTML.ML003 and ACTTB.TB002 = ACTML.ML004 and ACTTB.TB003 = ACTML.ML005
                            where ML009 like '%稅額調整%' and ML006 = '410202'
                         UNION ALL
                            select ML006 as 科目編號 
                            ,(select MA003 from ACTMA where MA001 = ACTML.ML001) as 科目名稱
                            ,SUBSTRING(ML002,1,6) as 傳票年月,ML003+'-'+ML004+' -'+ML005 as 傳票編號
                            ,ML009 as 摘要 ,TB012 as 備註
                            ,(case ML007 when '1' then ML008 else 0 end) as 本幣借方金額
                            ,(case ML007 when '-1' then ML008 else 0 end)  as 本幣貸方金額 
                            ,(case MA007 
	                            when '1' then ((case ML007 when '1' then ML008 else 0 end)-(case ML007 when '-1' then ML008 else 0 end)) 
	                            when '-1' then ((case ML007 when '-1' then ML008 else 0 end)-(case ML007 when '1' then ML008 else 0 end)) else 0 end ) as 貸借金額
                            ,'評價' as 項目,'材料成本' as 成本,SUBSTRING(ML009,CHARINDEX('#',ML009)+1,2) as 產品別
                            from ACTML
	                            left JOIN ACTMA on ACTMA.MA001 = ACTML.ML006
                                left JOIN ACTTB on ACTTB.TB001 = ACTML.ML003 and ACTTB.TB002 = ACTML.ML004 and ACTTB.TB003 = ACTML.ML005
                            where ML009 like '%存貨評價%' and ML006 = '510602'
                         UNION ALL
                            select ML006 as 科目編號 
                            ,(select MA003 from ACTMA where MA001 = ACTML.ML001) as 科目名稱
                            ,SUBSTRING(ML002,1,6) as 傳票年月,ML003+'-'+ML004+' -'+ML005 as 傳票編號
                            ,ML009 as 摘要 ,TB012 as 備註
                            ,(case ML007 when '1' then ML008 else 0 end) as 本幣借方金額
                            ,(case ML007 when '-1' then ML008 else 0 end)  as 本幣貸方金額 
                            ,(case MA007 
	                            when '1' then ((case ML007 when '1' then ML008 else 0 end)-(case ML007 when '-1' then ML008 else 0 end)) 
	                            when '-1' then ((case ML007 when '-1' then ML008 else 0 end)-(case ML007 when '1' then ML008 else 0 end)) else 0 end ) as 貸借金額
                            ,'下腳' as 項目,'材料成本' as 成本,'52' as 產品別
                            from ACTML
	                            left JOIN ACTMA on ACTMA.MA001 = ACTML.ML006
                                left JOIN ACTTB on ACTTB.TB001 = ACTML.ML003 and ACTTB.TB002 = ACTML.ML004 and ACTTB.TB003 = ACTML.ML005
                            where  ML006 = '510902'
                         UNION ALL
                            select ML006 as 科目編號 
                            ,(select MA003 from ACTMA where MA001 = ACTML.ML001) as 科目名稱
                            ,SUBSTRING(ML002,1,6) as 傳票年月,ML003+'-'+ML004+' -'+ML005 as 傳票編號
                            ,ML009 as 摘要 ,TB012 as 備註
                            ,(case ML007 when '1' then ML008 else 0 end) as 本幣借方金額
                            ,(case ML007 when '-1' then ML008 else 0 end)  as 本幣貸方金額 
                            ,(case MA007 
	                            when '1' then ((case ML007 when '1' then ML008 else 0 end)-(case ML007 when '-1' then ML008 else 0 end)) 
	                            when '-1' then ((case ML007 when '-1' then ML008 else 0 end)-(case ML007 when '1' then ML008 else 0 end)) else 0 end ) as 貸借金額
                            ,'成本調整' as 項目,'材料成本' as 成本,SUBSTRING(ML009,CHARINDEX('#',ML009)+1,2) as 產品別
                            from ACTML
	                            left JOIN ACTMA on ACTMA.MA001 = ACTML.ML006
                                left JOIN ACTTB on ACTTB.TB001 = ACTML.ML003 and ACTTB.TB002 = ACTML.ML004 and ACTTB.TB003 = ACTML.ML005
                            where ML009 like '%成本調整%' and ML006 in('510104','510204')
                            ) ACTML_ALL
                            where 傳票年月 between '{0}' and '{1}'
                            order by 傳票年月,科目編號", Date_S, Date_E);

            ////dt_ACRTB.Clear();   //銷貨單_勞務收入(佣金)
             sql_str_ACRTB = String.Format(
                @"select TG003 單據日期, left(TG003,6) 單據年月,TG004 客戶代號,MA002 客戶簡稱
                        ,TH001 銷貨單別,MQ002 單據名稱, TH002 銷貨單號, TB004 來源, MB006 產品別
                        , TH004 品號, TH027 結帳單別, TH028 結帳單號, TH029 結帳序號, TH037 本幣未稅金額
                    from COPTH
	                    left JOIN COPTG on COPTG.TG001=COPTH.TH001 and COPTG.TG002=COPTH.TH002
	                    left JOIN COPMA on COPMA.MA001=COPTG.TG004 
	                    left JOIN INVMB on INVMB.MB001=COPTH.TH004
	                    left JOIN CMSMQ on CMSMQ.MQ001  = COPTH.TH001 
	                    left JOIN ACRTB on ACRTB.TB001 = COPTH.TH027 and ACRTB.TB002 = COPTH.TH028 and ACRTB.TB003 = COPTH.TH029
                    where {0} and left(TG003,6) between '{1}' and '{2}' 
                    order by MB006,TG003", cond_ACRTB, Date_S, Date_E);
        }


        bool IsToForm1 = false; //紀錄是否要回到Form1
        protected override void OnClosing(CancelEventArgs e) //在視窗關閉時觸發
        {
            //DialogResult dr = MessageBox.Show("\"是\"回到主畫面 \r\n \"否\"關閉程式", "是否要關閉程式"
            //    , MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            //if (dr == DialogResult.Yes)
            //{
            IsToForm1 = true;
            //}
            //else if (dr == DialogResult.Cancel) 
            //{

            //}

            base.OnClosing(e);
            if (IsToForm1) //判斷是否要回到Form1
            {
                this.DialogResult = DialogResult.Yes; //利用DialogResult傳遞訊息
                fm_menu fm_menu = (fm_menu)this.Owner; //取得父視窗的參考
            }
            else
            {
                this.DialogResult = DialogResult.No;
            }
        }
    }
}
