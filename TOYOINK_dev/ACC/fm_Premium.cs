using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data;
using Myclass;
using ClosedXML.Excel;
using System.Globalization;

namespace TOYOINK_dev
{
    public partial class Premium : Form
    {
        /*20200203 加入出口模組，新增銷貨單別['2310','2311','2312']及21號新增客戶['AU-L6K']
         * 修改程式碼將單別及客戶改用宣告方式
         * 20200319 光阻 7%改為6%
         * 20200325 日本凸版 1%加入HSD
         * 20200515 更新公版路徑為\MIS自開發主檔\會計報表公版
         * 20201005 日本凸版權利金程式HKC-H4追加入1%計算，關閉 分頁跨月查詢
         * 20210401 日本凸版權利金程式 HKC-H5、CCPD 追加入1%計算
         */

        public MyClass MyCode;
        string str_enter = ((char)13).ToString() + ((char)10).ToString();

        DataTable dt_pAll = new DataTable();  //權利金-全部
        DataTable dt_dcAll = new DataTable();  //折讓-全部
        DataTable dt_pMPT = new DataTable();  //權利金-光阻
        DataTable dt_dcMPT = new DataTable(); //折讓-光阻
        DataTable dt_pJP = new DataTable();   //權利金-日本凸版
        DataTable dt_dcJP = new DataTable();  //折讓-日本凸板
        DataTable dt_hJP = new DataTable();  //手續費

        月曆 fm_月曆;
        string createday = DateTime.Now.ToString("yyyy/MM/dd");

        string save_as_JP = "", save_as_MPT = "", save_as_All = "", save_as_dcAll = "";

        DateTime date_s, date_s2;
        DateTime date_e, date_e2;

        string str_date_s, str_date_s2, str_date_s2_m, str_date_s2_15, str_date_s2_21;
        string str_date_e, str_date_e2, str_date_e2_m, str_date_e2_15, str_date_e2_21;

        string temp_excel;

        string copth_th001 = "'230', '2302', '230T', '234T', '235T','2310','2312','2320','2322'"
        , dcMPT_copti_ti001 = "'240','240T','241','241T'"
        , dcJP_copti_ti001 = "'240','241'";

        string coptg_tg004_15 = "'CHOT'"
        , coptg_tg004_21 = "'CPCF','CSOT','WCSOT','AU-L6K'";


        string defaultfilePath = "";

        public Premium()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();
            MyCode.strDbCon = "packet size=4096;user id=pwuser;password=sqlmis003;data source=192.168.128.219;persist security info=False;initial catalog=A01A;";
            //MyCode.strDbCon = "packet size=4096;user id=yj.chou;password=yjchou3369;data source=192.168.128.219;persist security info=False;initial catalog=Leader;";
            temp_excel = @"\\192.168.128.219\Conductor\Company\MIS自開發主檔\會計報表公版\權利金與折讓報表_temp.xlsx";
        }

        private void Premium_FormClosed(object sender, FormClosedEventArgs e)
        {
            Environment.Exit(Environment.ExitCode);
        }

        private void Premium_Load(object sender, EventArgs e)
        {
            tabControl2.SelectedIndex = 0;

            txt_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
            string filder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            txt_path.Text = filder;

            txt_date_s2.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_date_e2.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");

            //txt_date_s2.Text = "20190901";
            //txt_date_e2.Text = "20190930";
            string filder2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            txt_path2.Text = filder2;

            txterr.Text = "※權利金-光阻"+str_enter+
                          "　銷貨單：確認碼為\"Y\"、廠別為\"台南\" 002" + str_enter +
                          "　篩選單別為【'230', '2302', '230T', '234T', '235T','2310','2312','2320','2322'】" + str_enter +
                          "　抓取批號為【T開頭】" + str_enter +
                          "※權利金-日本凸版 上述條件，多加 扣除品號為'MPT'開頭"+str_enter +
                          "==========" + str_enter +
                          "※折讓-光阻"+ str_enter +
                         "　銷退單：" + str_enter +
                          "　篩選單別為【'240', '240T', '241', '241T'】" + str_enter +
                          "　抓取批號為【T開頭】" + str_enter +
                          "※折讓-日本凸版"+ str_enter +
                         "　銷退單：扣除品號為'MPT'開頭" + str_enter +
                          "　篩選單別為【'240','241'】" + str_enter +
                          "　抓取批號為【T開頭】" + str_enter +
                          "==========" + str_enter +
                          "※手續費 會計科目餘額表：科目為'623202'" + str_enter  +
                          "　摘要 前方為'客戶名稱'及篩選'帳款手續費'" + str_enter +
                          "==========" + str_enter +
                          "20201005 關閉分頁-跨月查詢";
            
            txterr2.Text = string.Format(
                @"一、CHOT    3月15日、6月15日、9月15日、12月15日以後的出貨，算在下一個月的營業額
15-31日列隔月營業額  ,4月.7月.10月.1月加回上個月扣除的營業額,重新計算

二、CPCF.CSOT.WCSOT.AU-L6K  3月21日、6月21日、9月21日、12月21日以後的出貨，算在下一個月的營業額
21-31日列隔月營業額  ,4月.7月.10月.1月加回上個月扣除的營業額,重新計算");
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
        }

        private void Btn_date_s2_Click(object sender, EventArgs e)
        {
            str_date_s2 = txt_date_s2.Text.Trim();
            this.fm_月曆 = new 月曆(this.txt_date_s2, this.Btn_date_s2, "單據起始日期");
        }

        private void Btn_date_e2_Click(object sender, EventArgs e)
        {
            str_date_e2 = txt_date_e2.Text.Trim();
            this.fm_月曆 = new 月曆(this.txt_date_e2, this.Btn_date_e2, "單據結束日期");
        }

        private void btn_down2_Click(object sender, EventArgs e)
        {

            date_s2 = DateTime.ParseExact(txt_date_s2.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e2 = DateTime.ParseExact(txt_date_e2.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s2.Text = DateTime.Parse(date_s2.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_date_e2.Text = DateTime.Parse(date_e2.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
        }

        private void btn_up2_Click(object sender, EventArgs e)
        {
            date_s2 = DateTime.ParseExact(txt_date_s2.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e2 = DateTime.ParseExact(txt_date_e2.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s2.Text = DateTime.Parse(date_s2.ToString("yyyy-MM-01")).AddMonths(1).ToString("yyyyMMdd");
            txt_date_e2.Text = DateTime.Parse(date_e2.ToString("yyyy-MM-01")).AddMonths(2).AddDays(-1).ToString("yyyyMMdd");
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

        private void txt_date_s2_TextChanged(object sender, EventArgs e)
        {
            Btn_acc2.Enabled = false;
            Btn_pre2.Enabled = false;
            Btn_dc2.Enabled = false;

            if (dgv_p_all.DataSource != null)
            {
                dt_pAll.Clear();
                dt_dcAll.Clear();
                dt_pMPT.Clear();
                dt_dcMPT.Clear();
                dt_pJP.Clear();
                dt_dcJP.Clear();
                dt_hJP.Clear();

                dgv_p_all.DataSource = null;
                dgv_dc_all.DataSource = null;
                dgv_p_MPT.DataSource = null;
                dgv_dc_MPT.DataSource = null;
                dgv_p_JP.DataSource = null;
                dgv_dc_JP.DataSource = null;
                dgv_hJP.DataSource = null;
            }
        }

        private void tabControl2_Selecting(object sender, TabControlCancelEventArgs e)
        {
            e.Cancel = true;
        }

        private void Btn_pd_Click(object sender, EventArgs e)
        {

        }

        private void txt_date_e2_TextChanged(object sender, EventArgs e)
        {
            Btn_acc2.Enabled = false;
            Btn_pre2.Enabled = false;
            Btn_dc2.Enabled = false;

            if (dgv_p_all.DataSource != null)
            {
                dt_pAll.Clear();
                dt_dcAll.Clear();
                dt_pMPT.Clear();
                dt_dcMPT.Clear();
                dt_pJP.Clear();
                dt_dcJP.Clear();
                dt_hJP.Clear();

                dgv_p_all.DataSource = null;
                dgv_dc_all.DataSource = null;
                dgv_p_MPT.DataSource = null;
                dgv_dc_MPT.DataSource = null;
                dgv_p_JP.DataSource = null;
                dgv_dc_JP.DataSource = null;
                dgv_hJP.DataSource = null;
            }
        }

        private void txt_date_s_TextChanged(object sender, EventArgs e)
        {
            Btn_acc.Enabled = false;
            Btn_pre.Enabled = false;
            Btn_dc.Enabled = false;

            if (dgv_p_all.DataSource != null)
            {
                dt_pAll.Clear();
                dt_dcAll.Clear();
                dt_pMPT.Clear();
                dt_dcMPT.Clear();
                dt_pJP.Clear();
                dt_dcJP.Clear();
                dt_hJP.Clear();

                dgv_p_all.DataSource = null;
                dgv_dc_all.DataSource = null;
                dgv_p_MPT.DataSource = null;
                dgv_dc_MPT.DataSource = null;
                dgv_p_JP.DataSource = null;
                dgv_dc_JP.DataSource = null;
                dgv_hJP.DataSource = null;
            }
        }

        private void txt_date_e_TextChanged(object sender, EventArgs e)
        {
            Btn_acc.Enabled = false;
            Btn_pre.Enabled = false;
            Btn_dc.Enabled = false;

            if (dgv_p_all.DataSource != null)
            {
                dt_pAll.Clear();
                dt_dcAll.Clear();
                dt_pMPT.Clear();
                dt_dcMPT.Clear();
                dt_pJP.Clear();
                dt_dcJP.Clear();
                dt_hJP.Clear();

                dgv_p_all.DataSource = null;
                dgv_dc_all.DataSource = null;
                dgv_p_MPT.DataSource = null;
                dgv_dc_MPT.DataSource = null;
                dgv_p_JP.DataSource = null;
                dgv_dc_JP.DataSource = null;
                dgv_hJP.DataSource = null;
            }
        }

        private void Btn_search_Click(object sender, EventArgs e)
        {
            Btn_acc.Enabled = false;
            Btn_pre.Enabled = false;
            Btn_dc.Enabled = false;

            date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            if (date_s > date_e)
            {
                MessageBox.Show("請修改日期區間", "日期格式錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dgv_p_all.DataSource != null)
            {
                dt_pAll.Clear();
                dt_dcAll.Clear();
                dt_pMPT.Clear();
                dt_dcMPT.Clear();
                dt_pJP.Clear();
                dt_dcJP.Clear();
                dt_hJP.Clear();

                dgv_p_all.DataSource = null;
                dgv_dc_all.DataSource = null;
                dgv_p_MPT.DataSource = null;
                dgv_dc_MPT.DataSource = null;
                dgv_p_JP.DataSource = null;
                dgv_dc_JP.DataSource = null;
                dgv_hJP.DataSource = null;
            }
            //dt_pAll.Clear();
            //dt_dcAll.Clear();
            //dt_pMPT.Clear();
            //dt_dcMPT.Clear();
            //dt_pJP.Clear();
            //dt_dcJP.Clear();
            //dt_hJP.Clear();

            //dgv_p_all.DataSource = null;
            //dgv_dc_all.DataSource = null;
            //dgv_p_MPT.DataSource = null;
            //dgv_dc_MPT.DataSource = null;
            //dgv_p_JP.DataSource = null;
            //dgv_dc_JP.DataSource = null;
            //dgv_hJP.DataSource = null;


            // 權利金-光阻 = 銷貨日報表全部
            // COPTH.TH020 = 'Y'    確認碼為"Y"
            // COPTG.TG010 = '002'  廠別為"台南"
            // COPTH.TH001 in ('230', '2302', '230T', '234T', '235T')  篩選單別【'230', '2302', '230T', '234T', '235T'】
            // COPTH.TH017 like 'T%'    抓取批號為【T開頭】
            string sql_str_pMPT = String.Format(
                @"SELECT COPTG.TG004 as 客戶代號,COPMA.MA002 as 客戶簡稱,COPTG.TG003 as 銷貨日期,COPTH.TH001 as 單別,
                COPTH.TH002 as 單身單號,COPTH.TH017 as 批號,COPTH.TH004 as 品號,COPTH.TH008 as 數量,
                COPTH.TH035 as 原幣未稅金額,COPTH.TH036 as 原幣稅額 ,COPTH.TH035 + COPTH.TH036 as 原幣合計金額,
                COPTG.TG011 as 幣別,COPTG.TG012 as 匯率,
                COPTH.TH037 as 本幣未稅金額,COPTH.TH038 as 本幣稅額,COPTH.TH037 + COPTH.TH038 as 本幣合計金額
                FROM COPTH as COPTH
                Inner JOIN COPTG as COPTG On COPTH.TH001 = COPTG.TG001 and COPTH.TH002 = COPTG.TG002
                Inner JOIN COPMA as COPMA On COPTG.TG004 = COPMA.MA001
                where COPTG.TG003 >= '{0}' and COPTG.TG003 <= '{1}' and COPTH.TH020 = 'Y' and COPTG.TG010 = '002'
                and COPTH.TH001 in ({2})
                and COPTH.TH017 like 'T%'
                ORDER BY COPTG.TG004 asc, COPTG.TG003 asc, COPTG.TG002 asc, COPTH.TH004 asc"
                , txt_date_s.Text.ToString().Trim(), txt_date_e.Text.ToString().Trim(),copth_th001);

            MyCode.Sql_dgv(sql_str_pMPT, dt_pMPT, dgv_p_MPT);
            MyCode.Sql_dgv(sql_str_pMPT, dt_pAll, dgv_p_all);

            // 權利金-光阻折讓
            //240 	折讓單-製品
            //240T 折讓單-製品 - 關係人
            //242     折讓單 - 商品
            //242T 折讓單-商品 - 關係人
            //批號為 T開頭，單別為[240.240T.241.241T]
            string sql_str_dcMPT = String.Format(
                @"SELECT COPTI.TI004 as 客戶,COPTG.TG003 as 銷貨日期,
                COPTI.TI001 as 折讓單別,COPTI.TI002 as 折讓單號,COPTC.TC012 as 客戶單號,
                COPTJ.TJ004 as 品號,COPTJ.TJ014 as 批號,COPTH.TH012 as 銷貨單價,COPTJ.TJ011 as 新單價,
                (COPTH.TH012 - COPTJ.TJ011) as 折讓差,COPTH.TH008 as 銷貨數量,
                COPTJ.TJ031 as 折讓金額,COPTJ.TJ032 as 折讓稅額,COPTJ.TJ033 as 台幣金額,COPTJ.TJ034 as 台幣稅額,COPTI.TI009 as 匯率,
                COPTJ.TJ033 + COPTJ.TJ034 as 台幣合計,COPTI.TI014 as 發票號碼,COPTI.TI006 as 廠別,COPTI.TI003 as 銷退日期
                FROM COPTJ as COPTJ
                Left JOIN COPTH as COPTH On COPTJ.TJ015 = COPTH.TH001 and COPTJ.TJ016 = COPTH.TH002 and COPTJ.TJ017 = COPTH.TH003
                Left JOIN COPTI as COPTI On COPTJ.TJ001 = COPTI.TI001 and COPTJ.TJ002 = COPTI.TI002
                Left JOIN COPTC as COPTC On COPTH.TH014 = COPTC.TC001 AND COPTH.TH015 = COPTC.TC002
                Left JOIN COPTG as COPTG On COPTH.TH001 = COPTG.TG001 AND COPTH.TH002 = COPTG.TG002
                where COPTI.TI003 >= '{0}' and COPTI.TI003 <= '{1}'
                and COPTJ.TJ014 like 'T%' and COPTI.TI001 in ({2})
                ORDER BY COPTI.TI003 asc, COPTI.TI004 asc, COPTG.TG003 asc"
                , txt_date_s.Text.ToString().Trim(), txt_date_e.Text.ToString().Trim(),dcMPT_copti_ti001);

            MyCode.Sql_dgv(sql_str_dcMPT, dt_dcMPT, dgv_dc_MPT);
            MyCode.Sql_dgv(sql_str_dcMPT, dt_dcAll, dgv_dc_all);


            // 權利金-日本凸版 
            // 同銷貨日報表條件，新增排除 品名為光阻 MPT 品項

            string sql_str_pJP = String.Format(
                @"SELECT COPTG.TG004 as 客戶代號,COPMA.MA002 as 客戶簡稱,COPTG.TG003 as 銷貨日期,COPTH.TH001 as 單別,
                COPTH.TH002 as 單身單號,COPTH.TH017 as 批號,COPTH.TH004 as 品號,COPTH.TH008 as 數量,
                COPTH.TH035 as 原幣未稅金額,COPTH.TH036 as 原幣稅額   ,COPTH.TH035 + COPTH.TH036 as 原幣合計金額,
                COPTG.TG011 as 幣別,COPTG.TG012 as 匯率,
                COPTH.TH037 as 本幣未稅金額,COPTH.TH038 as 本幣稅額,COPTH.TH037 + COPTH.TH038 as 本幣合計金額
                FROM COPTH as COPTH
                Inner JOIN COPTG as COPTG On COPTH.TH001 = COPTG.TG001 and COPTH.TH002 = COPTG.TG002
                Inner JOIN COPMA as COPMA On COPTG.TG004 = COPMA.MA001
                where COPTG.TG003 >= '{0}' and COPTG.TG003 <= '{1}' and COPTH.TH020 = 'Y' and COPTG.TG010 = '002'
                and COPTH.TH001 in ({2})
                and COPTH.TH017 like 'T%' and COPTH.TH004 not like 'MPT%'
                ORDER BY COPTG.TG004 asc, COPTG.TG003 asc, COPTG.TG002 asc, COPTH.TH004 asc"
                , txt_date_s.Text.ToString().Trim(), txt_date_e.Text.ToString().Trim(),copth_th001);


            MyCode.Sql_dgv(sql_str_pJP, dt_pJP, dgv_p_JP);

            // 權利金-日本凸版折讓
            //批號 T開頭，品號不等於'MPT'，單別為[240.241]
            string sql_str_dcJP = String.Format(
                @"SELECT COPTI.TI004 as 客戶,COPTG.TG003 as 銷貨日期,
               COPTI.TI001 as 折讓單別,COPTI.TI002 as 折讓單號,COPTC.TC012 as 客戶單號,
               COPTJ.TJ004 as 品號,COPTJ.TJ014 as 批號,COPTH.TH012 as 銷貨單價,COPTJ.TJ011 as 新單價,
               (COPTH.TH012 - COPTJ.TJ011) as 折讓差,COPTH.TH008 as 銷貨數量,
               COPTJ.TJ031 as 折讓金額,COPTJ.TJ032 as 折讓稅額,COPTJ.TJ033 as 台幣金額,COPTJ.TJ034 as 台幣稅額,COPTI.TI009 as 匯率,
               COPTJ.TJ033 + COPTJ.TJ034 as 台幣合計,COPTI.TI014 as 發票號碼,COPTI.TI006 as 廠別,COPTI.TI003 as 銷退日期
               FROM COPTJ as COPTJ
               Left JOIN COPTH as COPTH On COPTJ.TJ015 = COPTH.TH001 and COPTJ.TJ016 = COPTH.TH002 and COPTJ.TJ017 = COPTH.TH003
               Left JOIN COPTI as COPTI On COPTJ.TJ001 = COPTI.TI001 and COPTJ.TJ002 = COPTI.TI002
               Left JOIN COPTC as COPTC On COPTH.TH014 = COPTC.TC001 AND COPTH.TH015 = COPTC.TC002
               Left JOIN COPTG as COPTG On COPTH.TH001 = COPTG.TG001 AND COPTH.TH002 = COPTG.TG002
               where COPTI.TI003 >= '{0}' and COPTI.TI003 <= '{1}' and COPTJ.TJ014 like 'T%' and COPTH.TH004 not like 'MPT%'
               and COPTI.TI001 in ({2})
               ORDER BY COPTI.TI003 asc, COPTI.TI004 asc, COPTG.TG003 asc"
               , txt_date_s.Text.ToString().Trim(), txt_date_e.Text.ToString().Trim(),dcJP_copti_ti001);

            MyCode.Sql_dgv(sql_str_dcJP, dt_dcJP, dgv_dc_JP);

            //手續費
            //依摘要 篩選
            string sql_str_hJP = String.Format(
                @"select * from 
                    (select SUBSTRING( ML009 ,1, CHARINDEX (' ', ML009) -1) as 客戶名稱, 
                    ML001 as 科目編號 ,
                    (select MA003 from ACTMA where MA001 = ACTML.ML001) as 科目名稱,
                    ML002 as 傳票日期 ,
                    ML003+'-'+ML004+' -'+ML005 as 傳票編號,
                    ML009 as 摘要 ,
                    (case (SUBSTRING( ML009 ,1, CHARINDEX (' ', ML009) -1)) 
	                    when 'CSOT' then '2%'
	                    when 'WCSOT' then '2%'
	                    when 'CPCF' then '2%'
	                    when 'AU-L6K' then '2%'
	                    when 'AU-TC' then '2%'
	                    when 'AU-TY' then '2%'
	                    when 'BOE' then '1%'
	                    when 'CPD' then '1%'
	                    when 'HKC' then '1%'
	                    when 'HKC-H2' then '1%'
                        when 'HKC-H4' then '1%'
                        when 'HKC-H5' then '1%'
                        when 'CCPD' then '1%'
	                    when 'CHOT' then '1%'
                        when 'HSD' then '1%' end) as 手續費率,
                    (case ML007 when '1' then ML008 else 0 end) as 借方金額,
                    (case ML007 when '-1' then ML008 else 0 end)  as 貸方金額 ,
                    (case ML007 when '1' then '借餘' when '-1' then '貸餘' end) as 借貸 
                    from ACTML
                    where ML006 = '623202' and ML002 >='{0}' and ML002 <= '{1}' and ML009 like '%帳款手續費%') 手續費
                    where 手續費率 is not null
                    order by 手續費率 desc,客戶名稱,傳票日期"
                , txt_date_s.Text.ToString().Trim(), txt_date_e.Text.ToString().Trim());

            MyCode.Sql_dgv(sql_str_hJP, dt_hJP, dgv_hJP);

            tabControl1.SelectedIndex = 0;
            Btn_acc.Enabled = true;
            Btn_pre.Enabled = true;
            Btn_dc.Enabled = true;
        }



        private void Btn_file_Click(object sender, EventArgs e)
        {
            //if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    txt_path.Text = this.folderBrowserDialog1.SelectedPath;
            //}
            //else
            //{
            //    return;
            //}

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

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (dgv_p_all.DataSource != null)
            {
                //if (tabControl2.SelectedIndex == 1)
                //{

                //}

                if (Btn_acc.Enabled == true)
                {
                    Btn_acc.Enabled = false;
                    Btn_pre.Enabled = false;
                    Btn_dc.Enabled = false;
                }
                else if (Btn_acc2.Enabled == true) 
                {
                    Btn_acc2.Enabled = false;
                    Btn_pre2.Enabled = false;
                    Btn_dc2.Enabled = false;
                }

                dt_pAll.Clear();
                dt_dcAll.Clear();
                dt_pMPT.Clear();
                dt_dcMPT.Clear();
                dt_pJP.Clear();
                dt_dcJP.Clear();
                dt_hJP.Clear();

                dgv_p_all.DataSource = null;
                dgv_dc_all.DataSource = null;
                dgv_p_MPT.DataSource = null;
                dgv_dc_MPT.DataSource = null;
                dgv_p_JP.DataSource = null;
                dgv_dc_JP.DataSource = null;
                dgv_hJP.DataSource = null;
            }
        }

        private void btn_file2_Click(object sender, EventArgs e)
        {
            //if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    txt_path.Text = this.folderBrowserDialog1.SelectedPath;
            //}
            //else
            //{
            //    return;
            //}

            FolderBrowserDialog dialog2 = new FolderBrowserDialog();

            //首次defaultfilePath为空，按FolderBrowserDialog默认设置（即桌面）选择
            if (defaultfilePath != "")
            {
                //设置此次默认目录为上一次选中目录
                dialog2.SelectedPath = defaultfilePath;
            }

            if (dialog2.ShowDialog() == DialogResult.OK)
            {
                //记录选中的目录
                defaultfilePath = dialog2.SelectedPath;
                txt_path2.Text = defaultfilePath;
            }
        }

        private void Btn_acc_Click(object sender, EventArgs e)
        {
            if (dgv_p_MPT.DataSource is null)
            {
                MessageBox.Show("請先【查詢】", "無法轉出Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //保存成文件 權利金_光阻
            using (XLWorkbook wb_MPT = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel))
                {
                    var ws = templateWB.Worksheet(2);
                    var ws2 = templateWB.Worksheet(3);

                    ws.CopyTo(wb_MPT, "權利金_光阻");
                    ws2.CopyTo(wb_MPT, "折讓_光阻");
                }

                var wsheet_MPT = wb_MPT.Worksheet("權利金_光阻");
                var wsheet_dcMPT = wb_MPT.Worksheet("折讓_光阻");

                //== 折讓_光阻 ================================================================
                int rows_count_dcMPT = dt_dcMPT.Rows.Count;
                int j = 0; int x = 0;
                string cust_dcMPT = "";
                string sum_dcMPT = "";

                foreach (DataRow row in dt_dcMPT.Rows)
                {
                    if (cust_dcMPT.ToString() != "" && row[0].ToString() != cust_dcMPT.ToString())
                    {
                        wsheet_dcMPT.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_dcMPT.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_dcMPT.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_dcMPT.Cell(j + 7, 1).Value = cust_dcMPT;
                        wsheet_dcMPT.Cell(j + 7, 2).Value = "小計";
                        wsheet_dcMPT.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 7, 11).FormulaA1 = "=sum(K" + (x + 7) + ":K" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 12).FormulaA1 = "=sum(L" + (x + 7) + ":L" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 13).FormulaA1 = "=sum(M" + (x + 7) + ":M" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 14).FormulaA1 = "=sum(N" + (x + 7) + ":N" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 15).FormulaA1 = "=sum(O" + (x + 7) + ":O" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 7, 17).FormulaA1 = "=sum(Q" + (x + 7) + ":Q" + (j + 6) + ")";

                        sum_dcMPT += "N" + (j + 7) + "+";

                        x = j + 1;
                        j++;
                        //        cust = "1";
                    }

                    if (row[0].ToString() != cust_dcMPT.ToString())
                    {
                        wsheet_dcMPT.Cell(j + 7, 1).Value = row[0]; //客戶
                    }

                    wsheet_dcMPT.Cell(j + 7, 2).Value = row[1]; //銷貨日期
                    wsheet_dcMPT.Cell(j + 7, 3).Style.NumberFormat.Format = "@";
                    wsheet_dcMPT.Cell(j + 7, 3).Value = row[2]; //折讓單別
                    wsheet_dcMPT.Cell(j + 7, 4).Style.NumberFormat.Format = "@";
                    wsheet_dcMPT.Cell(j + 7, 4).Value = row[3]; //折讓單號
                    wsheet_dcMPT.Cell(j + 7, 5).Style.NumberFormat.Format = "@";
                    wsheet_dcMPT.Cell(j + 7, 5).Value = row[4]; //客戶單號
                    wsheet_dcMPT.Cell(j + 7, 6).Value = row[5]; //品號
                    wsheet_dcMPT.Cell(j + 7, 7).Value = row[6]; //批號
                    wsheet_dcMPT.Cell(j + 7, 8).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 8).Value = row[7]; //銷貨單價
                    wsheet_dcMPT.Cell(j + 7, 9).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 9).Value = row[8]; //新單價
                    wsheet_dcMPT.Cell(j + 7, 10).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 10).Value = row[9]; //折讓差

                    wsheet_dcMPT.Cell(j + 7, 11).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 11).Value = row[10]; //銷貨數量
                    wsheet_dcMPT.Cell(j + 7, 12).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 12).Value = row[11]; //折讓金額
                    wsheet_dcMPT.Cell(j + 7, 13).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 13).Value = row[12]; //折讓稅額
                    wsheet_dcMPT.Cell(j + 7, 14).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 14).Value = row[13]; //台幣金額
                    wsheet_dcMPT.Cell(j + 7, 15).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 15).Value = row[14]; //台幣稅額
                    wsheet_dcMPT.Cell(j + 7, 16).Style.NumberFormat.Format = "#,##0.0000";
                    wsheet_dcMPT.Cell(j + 7, 16).Value = row[15]; //匯率
                    wsheet_dcMPT.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 17).Value = row[16]; //台幣合計
                    wsheet_dcMPT.Cell(j + 7, 18).Value = row[17]; //發票號碼
                    wsheet_dcMPT.Cell(j + 7, 19).Value = row[18]; //廠別

                    wsheet_dcMPT.Cell(j + 7, 20).Value = row[19]; //銷退日

                    cust_dcMPT = row[0].ToString().Trim();

                    if ((rows_count_dcMPT - 1) == dt_dcMPT.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        j++;
                        wsheet_dcMPT.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_dcMPT.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_dcMPT.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_dcMPT.Cell(j + 7, 1).Value = cust_dcMPT;
                        wsheet_dcMPT.Cell(j + 7, 2).Value = "小計";
                        wsheet_dcMPT.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 7, 11).FormulaA1 = "=sum(K" + (x + 7) + ":K" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 12).FormulaA1 = "=sum(L" + (x + 7) + ":L" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 13).FormulaA1 = "=sum(M" + (x + 7) + ":M" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 14).FormulaA1 = "=sum(N" + (x + 7) + ":N" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 15).FormulaA1 = "=sum(O" + (x + 7) + ":O" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 7, 17).FormulaA1 = "=sum(Q" + (x + 7) + ":Q" + (j + 6) + ")";

                        sum_dcMPT += "N" + (j + 7);
                        x = j + 1;

                        j++;

                        wsheet_dcMPT.Cell("B4").Value = txt_date_s.Text.ToString() + " ~ " + txt_date_e.Text.ToString();
                        wsheet_dcMPT.Cell("B5").Value = createday;
                        wsheet_dcMPT.Cell(j + 7, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_dcMPT.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_dcMPT.Cell(j + 7, 3).Value = "總計";
                        wsheet_dcMPT.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 7, 11).FormulaA1 = "=SUMIF(B:B,\"小計\",K:K)";
                        wsheet_dcMPT.Cell(j + 7, 12).FormulaA1 = "=SUMIF(B:B,\"小計\",L:L)";
                        wsheet_dcMPT.Cell(j + 7, 13).FormulaA1 = "=SUMIF(B:B,\"小計\",M:M)";
                        wsheet_dcMPT.Cell(j + 7, 14).FormulaA1 = "=SUMIF(B:B,\"小計\",N:N)";
                        wsheet_dcMPT.Cell(j + 7, 15).FormulaA1 = "=SUMIF(B:B,\"小計\",O:O)";
                        wsheet_dcMPT.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 7, 17).FormulaA1 = "=SUMIF(B:B,\"小計\",Q:Q)";

                        j=j+2;
                        wsheet_dcMPT.Cell(j + 9, 13).Value = "UV製品-折讓";
                        wsheet_dcMPT.Cell(j + 10, 13).Value = " 製品-折讓";
                        wsheet_dcMPT.Cell(j + 10, 14).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 10, 14).FormulaA1 = "=SUMIF(B:B,\"小計\",N:N)";
                        wsheet_dcMPT.Cell(j + 11, 13).Value = " 製品-退回";
                        wsheet_dcMPT.Cell(j + 12, 13).Value = " 商品-折讓";
                        wsheet_dcMPT.Cell(j + 13, 13).Value = " 商品-退回";
                        wsheet_dcMPT.Cell(j + 14, 13).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_dcMPT.Range("M" + (j + 14) + ":N" + (j + 14)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_dcMPT.Cell(j + 14, 13).Value = "合計";
                        wsheet_dcMPT.Cell(j + 14, 14).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 14, 14).FormulaA1 = "=sum(N" + (j + 10) + ":N" + (j + 13) + ")";
                    }
                    j++;
                }

                //== 權利金-光阻 ================================================================
                //var worksheet = wb.Worksheets.Add("權利金_光阻");

                int rows_count_pMPT = dt_pMPT.Rows.Count;
                int i = 0; int y = 0;
                string cust_pMPT = "";
                string sum_pMPT = "";

                foreach (DataRow row in dt_pMPT.Rows)
                {
                    if (cust_pMPT.ToString() != "" && row[0].ToString() != cust_pMPT.ToString())
                    {
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_MPT.Cell(i + 5, 1).Value = cust_pMPT;
                        wsheet_MPT.Cell(i + 5, 3).Value = "小計";
                        wsheet_MPT.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_MPT.Cell(i + 5, 8).FormulaA1 = "=sum(H" + (y + 5) + ":H" + (i + 4) + ")";
                        wsheet_MPT.Range("N" + (i + 5) + ":P" + (i + 5)).Style.NumberFormat.Format = "#,##0";
                        wsheet_MPT.Cell(i + 5, 14).FormulaA1 = "=sum(N" + (y + 5) + ":N" + (i + 4) + ")";
                        wsheet_MPT.Cell(i + 5, 15).FormulaA1 = "=sum(O" + (y + 5) + ":O" + (i + 4) + ")";
                        wsheet_MPT.Cell(i + 5, 16).FormulaA1 = "=sum(P" + (y + 5) + ":P" + (i + 4) + ")";

                        y = i + 1;
                        i++;

                    }

                    //填入excel欄位值
                    wsheet_MPT.Cell(i + 5, 1).Value = row[0]; //客戶代號
                    wsheet_MPT.Cell(i + 5, 2).Value = row[1]; //客戶簡稱
                    wsheet_MPT.Cell(i + 5, 3).Value = row[2]; //銷貨日期
                    wsheet_MPT.Cell(i + 5, 4).Style.NumberFormat.Format = "@";
                    wsheet_MPT.Cell(i + 5, 4).Value = row[3]; //單別
                    wsheet_MPT.Cell(i + 5, 5).Style.NumberFormat.Format = "@";
                    wsheet_MPT.Cell(i + 5, 5).Value = row[4]; //單號
                    wsheet_MPT.Cell(i + 5, 6).Value = row[5]; //批號
                    wsheet_MPT.Cell(i + 5, 7).Value = row[6]; //品號

                    if (row[6].ToString().Substring(0, 3) == "MPT")
                    {
                        sum_pMPT += "N" + (i + 5) + "+";
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDD7EE");
                    }

                    wsheet_MPT.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                    wsheet_MPT.Cell(i + 5, 8).Value = row[7]; //數量
                    wsheet_MPT.Cell(i + 5, 9).Style.NumberFormat.Format = "#,##0";
                    wsheet_MPT.Cell(i + 5, 9).Value = row[8]; //原幣未稅金額
                    wsheet_MPT.Cell(i + 5, 10).Style.NumberFormat.Format = "#,##0";
                    wsheet_MPT.Cell(i + 5, 10).Value = row[9]; //原幣稅額
                    wsheet_MPT.Cell(i + 5, 11).Style.NumberFormat.Format = "#,##0";
                    wsheet_MPT.Cell(i + 5, 11).Value = row[10]; //原幣合計金額
                    wsheet_MPT.Cell(i + 5, 12).Value = row[11]; //幣別
                    wsheet_MPT.Cell(i + 5, 13).Style.NumberFormat.Format = "#,##0.000";
                    wsheet_MPT.Cell(i + 5, 13).Value = row[12]; //匯率
                    wsheet_MPT.Cell(i + 5, 14).Style.NumberFormat.Format = "#,##0";
                    wsheet_MPT.Cell(i + 5, 14).Value = row[13]; //本幣未稅金額
                    wsheet_MPT.Cell(i + 5, 15).Style.NumberFormat.Format = "#,##0";
                    wsheet_MPT.Cell(i + 5, 15).Value = row[14]; //本幣稅額
                    wsheet_MPT.Cell(i + 5, 16).Style.NumberFormat.Format = "#,##0";
                    wsheet_MPT.Cell(i + 5, 16).Value = row[15]; //本幣合計金額

                    cust_pMPT = row[0].ToString().Trim();

                    if ((rows_count_pMPT - 1) == dt_pMPT.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        i++;
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_MPT.Cell(i + 5, 1).Value = cust_pMPT;
                        wsheet_MPT.Cell(i + 5, 3).Value = "小計";
                        wsheet_MPT.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_MPT.Cell(i + 5, 8).FormulaA1 = "=sum(H" + (y + 5) + ":H" + (i + 4) + ")";
                        wsheet_MPT.Range("N" + (i + 5) + ":P" + (i + 5)).Style.NumberFormat.Format = "#,##0";
                        wsheet_MPT.Cell(i + 5, 14).FormulaA1 = "=sum(N" + (y + 5) + ":N" + (i + 4) + ")";
                        wsheet_MPT.Cell(i + 5, 15).FormulaA1 = "=sum(O" + (y + 5) + ":O" + (i + 4) + ")";
                        wsheet_MPT.Cell(i + 5, 16).FormulaA1 = "=sum(P" + (y + 5) + ":P" + (i + 4) + ")";

                        y = i + 1;

                        wsheet_MPT.Cell("B2").Value = txt_date_s.Text.ToString() + "~" + txt_date_e.Text.ToString();
                        wsheet_MPT.Cell("B3").Value = createday;
                        wsheet_MPT.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_MPT.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_MPT.Cell(i + 6, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_MPT.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_MPT.Cell(i + 6, 3).Value = "總計";
                        wsheet_MPT.Cell(i + 6, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_MPT.Cell(i + 6, 8).FormulaA1 = "=SUMIF(C:C,\"小計\",H:H)";
                        wsheet_MPT.Range("N" + (i + 6) + ":P" + (i + 6)).Style.NumberFormat.Format = "#,##0";
                        wsheet_MPT.Cell(i + 6, 14).FormulaA1 = "=SUMIF(C:C,\"小計\",N:N)";
                        wsheet_MPT.Cell(i + 6, 15).FormulaA1 = "=SUMIF(C:C,\"小計\",O:O)";
                        wsheet_MPT.Cell(i + 6, 16).FormulaA1 = "=SUMIF(C:C,\"小計\",P:P)";

                        wsheet_MPT.Cell(i + 10, 13).Value = "折讓";
                        wsheet_MPT.Cell(i + 10, 14).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_MPT.Cell(i + 10, 14).FormulaA1 = "=-折讓_光阻!N" + (j + 13);

                        wsheet_MPT.Range("M" + (i + 11) + ":N" + (i + 11)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_MPT.Cell(i + 11, 13).Value = "合計";
                        wsheet_MPT.Cell(i + 11, 14).Style.NumberFormat.Format = "#,##0";
                        wsheet_MPT.Cell(i + 11, 14).FormulaA1 = "=N" + (i + 6) + "+N" + (i + 10);

                        wsheet_MPT.Range("F" + (i + 14) + ":N" + (i + 17)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_MPT.Range("F" + (i + 14) + ":N" + (i + 17)).Style.Font.FontSize = 14;
                        wsheet_MPT.Range("K" + (i + 15) + ":N" + (i + 17)).Style.NumberFormat.Format = "#,##0";


                        wsheet_MPT.Range("F" + (i + 14) + ":N" + (i + 17)).Style.Border.OutsideBorder = XLBorderStyleValues.Double;

                        wsheet_MPT.Range("H" + (i + 14) + ":H" + (i + 16)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsheet_MPT.Cell(i + 14, 11).Value = "銷貨";
                        wsheet_MPT.Cell(i + 14, 12).Value = "6%權利金";
                        wsheet_MPT.Cell(i + 14, 13).Value = "10%稅金";
                        wsheet_MPT.Cell(i + 14, 14).Value = "淨額";

                        wsheet_MPT.Cell(i + 15, 6).Value = "Royalty";
                        wsheet_MPT.Cell(i + 15, 8).Value = "光阻";
                        wsheet_MPT.Cell(i + 15, 10).Value = "6%";
                        wsheet_MPT.Cell(i + 15, 11).FormulaA1 = "N" + (i + 6) + "+N" + (i + 10) + "-K" + (i + 16);
                        wsheet_MPT.Cell(i + 15, 12).FormulaA1 = "= ROUND(K" + (i + 15) + "* 0.06, 0)";
                        wsheet_MPT.Cell(i + 15, 13).FormulaA1 = "= ROUND(L" + (i + 15) + "* 0.1, 0)";
                        wsheet_MPT.Cell(i + 15, 14).FormulaA1 = "= L" + (i + 15) + "-M" + (i + 15);

                        wsheet_MPT.Cell(i + 16, 8).Value = "MPT";
                        wsheet_MPT.Cell(i + 16, 10).Value = "2%";
                        
                        if (sum_pMPT.Length != 0) { 
                        wsheet_MPT.Cell(i + 16, 11).FormulaA1 = sum_pMPT.Substring(0, sum_pMPT.Length - 1);
                        }
                        wsheet_MPT.Cell(i + 16, 12).FormulaA1 = "= ROUND(K" + (i + 16) + "* 0.02, 0)";
                        wsheet_MPT.Cell(i + 16, 13).FormulaA1 = "= ROUND(L" + (i + 16) + "* 0.1, 0)";
                        wsheet_MPT.Cell(i + 16, 14).FormulaA1 = "= L" + (i + 16) + "-M" + (i + 16);

                        wsheet_MPT.Cell(i + 17, 8).Value = "光阻+MPT";
                        wsheet_MPT.Cell(i + 17, 11).FormulaA1 = "=sum(K" + (i + 15) + ":K" + (i + 16) + ")";
                        wsheet_MPT.Cell(i + 17, 12).FormulaA1 = "=sum(L" + (i + 15) + ":L" + (i + 16) + ")";
                        wsheet_MPT.Cell(i + 17, 13).FormulaA1 = "=sum(M" + (i + 15) + ":M" + (i + 16) + ")";
                        wsheet_MPT.Cell(i + 17, 14).FormulaA1 = "=sum(N" + (i + 15) + ":N" + (i + 16) + ")";

                    }
                    i++;
                }
                //worksheet.Columns().AdjustToContents();
                //worksheet2.Columns().AdjustToContents();

                wsheet_MPT.Position = 1;
                wsheet_MPT.Column("K").Width = 16;
                wsheet_MPT.Column("L").Width = 16;
                wsheet_MPT.Column("M").Width = 16;
                wsheet_MPT.Column("N").Width = 16;

                save_as_MPT = txt_path.Text.ToString().Trim() + @"\\銷貨明細(權利金)光阻" + txt_date_e.Text.ToString().Substring(0, 6) + "_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";

                wb_MPT.SaveAs(save_as_MPT);

                //打开文件
                //System.Diagnostics.Process.Start(save_as_MPT);
            }

            //保存成文件 日本凸版
            using (XLWorkbook wb_JP = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel))
                {
                    var ws = templateWB.Worksheet(1);
                    var ws2 = templateWB.Worksheet(3);
                    var ws3 = templateWB.Worksheet(4);

                    ws.CopyTo(wb_JP, "權利金_日本凸版");
                    ws2.CopyTo(wb_JP, "折讓_日本凸版");
                    ws3.CopyTo(wb_JP, "手續費");
                }

                //var worksheet2 = wb_JP.Worksheets.Add("折讓_光阻");
                var wsheet_JP = wb_JP.Worksheet("權利金_日本凸版");
                var wsheet_dcJP = wb_JP.Worksheet("折讓_日本凸版");
                var wsheet_hJP = wb_JP.Worksheet("手續費");

                //== 手續費-日本凸版 ================================================================
                int rows_count_hJP = dt_hJP.Rows.Count;
                int k = 0; int z = 0; int q = 0; int r = 0;
                string cust_hJP = "";
                string custid_hJP = "";
                int[] custidnum_hJP = new int[20];
                string[] custidname_hJP = new string[20];
                int[] num_hJP = new int[20];

                if (rows_count_hJP == 0)
                {
                    wsheet_hJP.Cell("B2").Value = txt_date_s.Text.ToString() + "~" + txt_date_e.Text.ToString();
                    wsheet_hJP.Cell("B3").Value = createday;

                    wsheet_hJP.Cell(k + 9, 7).Value = "手續費";

                    wsheet_hJP.Range("G" + (k + 10) + ":H" + (k + 10)).Style.Fill.BackgroundColor = XLColor.Yellow;
                    wsheet_hJP.Cell(k + 10, 7).Value = "合計";
                    wsheet_hJP.Cell(k + 10, 8).Value = "0";

                    wsheet_hJP.Range("F" + (k + 13) + ":H" + (k + 15)).Style.Fill.BackgroundColor = XLColor.Yellow;
                    wsheet_hJP.Range("H" + (k + 13) + ":H" + (k + 15)).Style.NumberFormat.Format = "#,##0";
                    wsheet_hJP.Range("F" + (k + 13) + ":F" + (k + 14)).Style.Font.FontSize = 10;


                    wsheet_hJP.Range("H" + (k + 13) + ":H" + (k + 15)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    wsheet_hJP.Range("F" + (k + 14) + ":H" + (k + 14)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    wsheet_hJP.Range("F" + (k + 15) + ":H" + (k + 15)).Style.Border.BottomBorder = XLBorderStyleValues.Double;

                    wsheet_hJP.Cell(k + 13, 6).Value = "手續費(2%)-CSOT、WCSOT、CPCF、AU-L6K、AU-TC、AU-TY";
                    wsheet_hJP.Cell(k + 13, 8).FormulaA1 = "=SUM(SUMIFS(H:H,A:A,{\"CSOT\",\"WCSOT\",\"CPCF\",\"AU-L6K\",\"AU-TC\",\"AU-TY\"},B:B,\"小計\"))";
                    num_hJP[0] = k + 13;

                    //TODO:20200325 日本凸版加入HSD、20201005 加入HKC-H4、20210401 加入HKC-H5、CCPD
                    wsheet_hJP.Cell(k + 14, 6).Value = "手續費(1%)-BOE、CPD、HKC、HKC-H2、HKC-H4、HKC-H5、CHOT、HSD、CCPD";
                    wsheet_hJP.Cell(k + 14, 8).FormulaA1 = "=SUM(SUMIFS(H:H,A:A,{\"BOE\",\"CPD\",\"HKC\",\"HKC-H2\",\"HKC-H4\",\"HKC-H5\",\"CHOT\",\"HSD\",\"HCCPD\"},B:B,\"小計\"))";
                    num_hJP[1] = k + 14;

                    wsheet_hJP.Cell(k + 15, 6).Value = "合計";
                    wsheet_hJP.Cell(k + 15, 8).FormulaA1 = "=sum(H" + (k + 13) + ":H" + (k + 14) + ")";
                }
                else
                {
                    foreach (DataRow row in dt_hJP.Rows)
                    {
                        if (cust_hJP.ToString() != "" && row[0].ToString() != cust_hJP.ToString())
                        {
                            wsheet_hJP.Range("A" + (k + 5) + ":J" + (k + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Range("A" + (k + 5) + ":J" + (k + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Range("A" + (k + 5) + ":J" + (k + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                            wsheet_hJP.Cell(k + 5, 1).Value = custid_hJP;
                            wsheet_hJP.Cell(k + 5, 2).Value = "小計";
                            wsheet_hJP.Cell(k + 5, 8).FormulaA1 = "=sum(H" + (z + 5) + ":H" + (k + 4) + ")";

                            custidname_hJP[q] = custid_hJP;

                            q += 1;
                            z = k + 1;

                            k++;
                        }

                        wsheet_hJP.Cell(k + 5, 1).Value = row[0]; //客戶名稱
                        wsheet_hJP.Cell(k + 5, 2).Value = row[1]; //科目編號
                        wsheet_hJP.Cell(k + 5, 3).Value = row[2]; //科目名稱
                        wsheet_hJP.Cell(k + 5, 4).Value = row[3]; //傳票日期
                        wsheet_hJP.Cell(k + 5, 5).Value = row[4]; //傳票編號
                        wsheet_hJP.Cell(k + 5, 6).Value = row[5]; //摘要
                        wsheet_hJP.Cell(k + 5, 7).Value = row[6]; //手續費率
                        wsheet_hJP.Cell(k + 5, 8).Style.NumberFormat.Format = "#,##0";
                        wsheet_hJP.Cell(k + 5, 8).Value = row[7]; //借方金額
                        wsheet_hJP.Cell(k + 5, 9).Style.NumberFormat.Format = "#,##0";
                        wsheet_hJP.Cell(k + 5, 9).Value = row[8]; //貸方金額
                        wsheet_hJP.Cell(k + 5, 10).Value = row[9]; //借貸

                        cust_hJP = row[0].ToString().Trim();
                        custid_hJP = row[0].ToString().Trim();

                        if ((rows_count_hJP - 1) == dt_hJP.Rows.IndexOf(row)) //資料列結尾運算
                        {
                            k++;
                            wsheet_hJP.Range("A" + (k + 5) + ":J" + (k + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Range("A" + (k + 5) + ":J" + (k + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Range("A" + (k + 5) + ":J" + (k + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                            wsheet_hJP.Cell(k + 5, 1).Value = custid_hJP;
                            wsheet_hJP.Cell(k + 5, 2).Value = "小計";
                            wsheet_hJP.Cell(k + 5, 8).FormulaA1 = "=sum(H" + (z + 5) + ":H" + (k + 4) + ")";

                            custidname_hJP[q] = custid_hJP;
                            //custidnum_dcJP[n] = (k + 5);

                            //sum_dcJP += "N" + (k + 5);
                            //x = j + 1;

                            z++;

                            wsheet_hJP.Cell("B2").Value = txt_date_s.Text.ToString() + "~" + txt_date_e.Text.ToString();
                            wsheet_hJP.Cell("B3").Value = createday;
                            wsheet_hJP.Range("A" + (k + 6) + ":J" + (k + 6)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Range("A" + (k + 6) + ":J" + (k + 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Cell(k + 6, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                            wsheet_hJP.Range("A" + (k + 6) + ":J" + (k + 6)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                            wsheet_hJP.Cell(k + 6, 2).Value = "總計";
                            wsheet_hJP.Cell(k + 6, 8).Style.NumberFormat.Format = "#,##0.00";
                            wsheet_hJP.Cell(k + 6, 8).FormulaA1 = "=SUMIF(B:B,\"小計\",H:H)";

                            wsheet_hJP.Cell(k + 9, 7).Value = "手續費";

                            for (int num = 0; num <= q; num++)
                            {
                                wsheet_hJP.Cell(k + 10, 7).Value = custidname_hJP[num];
                                wsheet_hJP.Cell(k + 10, 8).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                                wsheet_hJP.Cell(k + 10, 8).FormulaA1 = "=SUMIFS(H:H,A:A,\"" + custidname_hJP[num] + "\",B:B,\"小計\")";
                                k++;
                                r++;
                            }

                            wsheet_hJP.Cell(k + 10, 7).Value = "合計";
                            wsheet_hJP.Cell(k + 10, 8).Style.NumberFormat.Format = "#,##0";

                            wsheet_hJP.Range("G" + (k + 10) + ":H" + (k + 10)).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wsheet_hJP.Cell(k + 10, 8).FormulaA1 = "=SUM(H" + (k + 10 - r) + ":H" + (k + 9) + ")";

                            wsheet_hJP.Range("F" + (k + 13) + ":H" + (k + 15)).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wsheet_hJP.Range("H" + (k + 13) + ":H" + (k + 15)).Style.NumberFormat.Format = "#,##0";
                            wsheet_hJP.Range("F" + (k + 13) + ":F" + (k + 14)).Style.Font.FontSize = 10;


                            wsheet_hJP.Range("H" + (k + 13) + ":H" + (k + 15)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            wsheet_hJP.Range("F" + (k + 14) + ":H" + (k + 14)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Range("F" + (k + 15) + ":H" + (k + 15)).Style.Border.BottomBorder = XLBorderStyleValues.Double;

                            wsheet_hJP.Cell(k + 13, 6).Value = "手續費(2%)-CSOT、WCSOT、CPCF、AU-L6K、AU-TC、AU-TY";
                            wsheet_hJP.Cell(k + 13, 8).FormulaA1 = "=SUM(SUMIFS(H:H,A:A,{\"CSOT\",\"WCSOT\",\"CPCF\",\"AU-L6K\",\"AU-TC\",\"AU-TY\"},B:B,\"小計\"))";
                            num_hJP[0] = k + 13;

                            //TODO:20200325 日本凸版加入HSD、20201005 加入HKC-H4、20210401 加入、20201005 加入HKC-H4
                            wsheet_hJP.Cell(k + 14, 6).Value = "手續費(1%)-BOE、CPD、HKC、HKC-H2、HKC-H4、HKC-H5、CHOT、HSD、CCPD";
                            wsheet_hJP.Cell(k + 14, 8).FormulaA1 = "=SUM(SUMIFS(H:H,A:A,{\"BOE\",\"CPD\",\"HKC\",\"HKC-H2\",\"HKC-H4\",\"HKC-H5\",\"CHOT\",\"HSD\",\"CCPD\"},B:B,\"小計\"))";
                            num_hJP[1] = k + 14;

                            wsheet_hJP.Cell(k + 15, 6).Value = "合計";
                            wsheet_hJP.Cell(k + 15, 8).FormulaA1 = "=sum(H" + (k + 13) + ":H" + (k + 14) + ")";
                        }
                        k++;
                    }
                }

                //== 折讓-日本凸版 ================================================================
                int rows_count_dcJP = dt_dcJP.Rows.Count;
                int j = 0; int x = 0; int n = 0;
                string cust_dcJP = "";
                string custid_dcJP = "";
                int[] custidnum_dcJP = new int[20];
                string[] custidname_dcJP = new string[20];
                int[] num_dcJP = new int[20];

                foreach (DataRow row in dt_dcJP.Rows)
                {
                    if (cust_dcJP.ToString() != "" && row[0].ToString() != cust_dcJP.ToString())
                    {
                        wsheet_dcJP.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_dcJP.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_dcJP.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_dcJP.Cell(j + 7, 1).Value = custid_dcJP;
                        wsheet_dcJP.Cell(j + 7, 2).Value = "小計";
                        wsheet_dcJP.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 7, 11).FormulaA1 = "=sum(K" + (x + 7) + ":K" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 12).FormulaA1 = "=sum(L" + (x + 7) + ":L" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 13).FormulaA1 = "=sum(M" + (x + 7) + ":M" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 14).FormulaA1 = "=sum(N" + (x + 7) + ":N" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 15).FormulaA1 = "=sum(O" + (x + 7) + ":O" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 7, 17).FormulaA1 = "=sum(Q" + (x + 7) + ":Q" + (j + 6) + ")";

                        switch (custid_dcJP)
                        {
                            case "CSOT":
                            case "WCSOT":
                            case "CPCF":
                            case "AU-L6K":
                            case "AU-TC":
                            case "AU-TY":
                            case "BOE":
                            case "CPD":
                            case "HKC":
                            case "HKC-H2":
                            case "HKC-H4":
                            case "HKC-H5":
                            case "CCPD":
                            case "CHOT":
                            case "HSD":
                                custidname_dcJP[n] = custid_dcJP;
                                custidnum_dcJP[n] = (j + 7);
                                break;
                            default:
                                break;
                        }

                        //sum_dcJP += "N" + (j + 7) + "+";

                        n += 1;
                        x = j + 1;
                        j++;
                    }

                    if (row[0].ToString() != cust_dcJP.ToString())
                    {
                        wsheet_dcJP.Cell(j + 7, 1).Value = row[0]; //客戶
                    }

                    wsheet_dcJP.Cell(j + 7, 2).Value = row[1]; //銷貨日期
                    wsheet_dcJP.Cell(j + 7, 3).Style.NumberFormat.Format = "@";
                    wsheet_dcJP.Cell(j + 7, 3).Value = row[2]; //折讓單別
                    wsheet_dcJP.Cell(j + 7, 4).Style.NumberFormat.Format = "@";
                    wsheet_dcJP.Cell(j + 7, 4).Value = row[3]; //折讓單號
                    wsheet_dcJP.Cell(j + 7, 5).Style.NumberFormat.Format = "@";
                    wsheet_dcJP.Cell(j + 7, 5).Value = row[4]; //客戶單號
                    wsheet_dcJP.Cell(j + 7, 6).Value = row[5]; //品號
                    wsheet_dcJP.Cell(j + 7, 7).Value = row[6]; //批號
                    wsheet_dcJP.Cell(j + 7, 8).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 8).Value = row[7]; //銷貨單價
                    wsheet_dcJP.Cell(j + 7, 9).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 9).Value = row[8]; //新單價
                    wsheet_dcJP.Cell(j + 7, 10).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 10).Value = row[9]; //折讓差

                    wsheet_dcJP.Cell(j + 7, 11).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 11).Value = row[10]; //銷貨數量
                    wsheet_dcJP.Cell(j + 7, 12).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 12).Value = row[11]; //折讓金額
                    wsheet_dcJP.Cell(j + 7, 13).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 13).Value = row[12]; //折讓稅額
                    wsheet_dcJP.Cell(j + 7, 14).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 14).Value = row[13]; //台幣金額
                    wsheet_dcJP.Cell(j + 7, 15).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 15).Value = row[14]; //台幣稅額
                    wsheet_dcJP.Cell(j + 7, 16).Style.NumberFormat.Format = "#,##0.0000";
                    wsheet_dcJP.Cell(j + 7, 16).Value = row[15]; //匯率
                    wsheet_dcJP.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 17).Value = row[16]; //台幣合計
                    wsheet_dcJP.Cell(j + 7, 18).Value = row[17]; //發票號碼
                    wsheet_dcJP.Cell(j + 7, 19).Value = row[18]; //廠別

                    wsheet_dcJP.Cell(j + 7, 20).Value = row[19]; //銷退日

                    cust_dcJP = row[0].ToString().Trim();
                    custid_dcJP = row[0].ToString().Trim();

                    if ((rows_count_dcJP - 1) == dt_dcJP.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        j++;
                        wsheet_dcJP.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_dcJP.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_dcJP.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_dcJP.Cell(j + 7, 1).Value = custid_dcJP;
                        wsheet_dcJP.Cell(j + 7, 2).Value = "小計";
                        wsheet_dcJP.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 7, 11).FormulaA1 = "=sum(K" + (x + 7) + ":K" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 12).FormulaA1 = "=sum(L" + (x + 7) + ":L" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 13).FormulaA1 = "=sum(M" + (x + 7) + ":M" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 14).FormulaA1 = "=sum(N" + (x + 7) + ":N" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 15).FormulaA1 = "=sum(O" + (x + 7) + ":O" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 7, 17).FormulaA1 = "=sum(Q" + (x + 7) + ":Q" + (j + 6) + ")";

                        custidname_dcJP[n] = custid_dcJP;
                        custidnum_dcJP[n] = (j + 7);

                        //sum_dcJP += "N" + (j + 7);
                        x = j + 1;

                        j++;

                        wsheet_dcJP.Cell("B4").Value = txt_date_s.Text.ToString() + " ~ " + txt_date_e.Text.ToString();
                        wsheet_dcJP.Cell("B5").Value = createday;

                        wsheet_dcJP.Cell(j + 7, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_dcJP.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_dcJP.Cell(j + 7, 3).Value = "總計";
                        wsheet_dcJP.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 7, 11).FormulaA1 = "=SUMIF(B:B,\"小計\",K:K)";
                        wsheet_dcJP.Cell(j + 7, 12).FormulaA1 = "=SUMIF(B:B,\"小計\",L:L)";
                        wsheet_dcJP.Cell(j + 7, 13).FormulaA1 = "=SUMIF(B:B,\"小計\",M:M)";
                        wsheet_dcJP.Cell(j + 7, 14).FormulaA1 = "=SUMIF(B:B,\"小計\",N:N)";
                        wsheet_dcJP.Cell(j + 7, 15).FormulaA1 = "=SUMIF(B:B,\"小計\",O:O)";
                        wsheet_dcJP.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 7, 17).FormulaA1 = "=SUMIF(B:B,\"小計\",Q:Q)";

                        j = j + 2;
                        wsheet_dcJP.Range("N" + (j + 10) + ":N" + (j + 15)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 9, 13).Value = "UV製品-折讓";
                        wsheet_dcJP.Cell(j + 10, 13).Value = " 製品-折讓2%";
                        wsheet_dcJP.Cell(j + 10, 14).FormulaA1 = "=SUM(SUMIFS(N:N,A:A,{\"CSOT\",\"WCSOT\",\"CPCF\",\"AU-L6K\",\"AU-TC\",\"AU-TY\"},B:B,\"小計\"))";
                        num_dcJP[0] = (j + 10);

                        wsheet_dcJP.Cell(j + 11, 13).Value = " 製品-折讓1%";
                        wsheet_dcJP.Cell(j + 11, 14).FormulaA1 = "=SUM(SUMIFS(N:N,A:A,{\"BOE\",\"CPD\",\"HKC\",\"HKC-H2\",\"HKC-H4\",\"HKC-H5\",\"CHOT\",\"HSD\",\"CCPD\"},B:B,\"小計\"))";
                        num_dcJP[1] = (j + 11);

                        wsheet_dcJP.Cell(j + 12, 13).Value = " 製品-退回";
                        wsheet_dcJP.Cell(j + 13, 13).Value = " 商品-折讓";
                        wsheet_dcJP.Cell(j + 14, 13).Value = " 商品-退回";
                        wsheet_dcJP.Cell(j + 15, 13).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_dcJP.Range("M" + (j + 15) + ":N" + (j + 15)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_dcJP.Cell(j + 15, 13).Value = "合計";
                        wsheet_dcJP.Cell(j + 15, 14).FormulaA1 = "=sum(N" + (j + 10) + ":N" + (j + 14) + ")";
                    }
                    j++;
                }

                //== 權利金-日本凸版 ================================================================
                int rows_count_pJP = dt_pJP.Rows.Count;
                int i = 0; int y = 0; int m = 0; int p = 0;
                string cust_pJP = "";
                string custid_pJP = "";
                string[] custidname_pJP = new string[20];
                int[] sumrow_pJP = new int[20];

                foreach (DataRow row in dt_pJP.Rows)
                {
                    if (cust_pJP.ToString() != "" && row[0].ToString() != cust_pJP.ToString())
                    {
                        wsheet_JP.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_JP.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_JP.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_JP.Cell(i + 5, 1).Value = custid_pJP;
                        wsheet_JP.Cell(i + 5, 3).Value = "小計";
                        wsheet_JP.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_JP.Cell(i + 5, 8).FormulaA1 = "=sum(H" + (y + 5) + ":H" + (i + 4) + ")";
                        wsheet_JP.Range("N" + (i + 5) + ":P" + (i + 5)).Style.NumberFormat.Format = "#,##0";
                        wsheet_JP.Cell(i + 5, 14).FormulaA1 = "=sum(N" + (y + 5) + ":N" + (i + 4) + ")";
                        wsheet_JP.Cell(i + 5, 15).FormulaA1 = "=sum(O" + (y + 5) + ":O" + (i + 4) + ")";
                        wsheet_JP.Cell(i + 5, 16).FormulaA1 = "=sum(P" + (y + 5) + ":P" + (i + 4) + ")";

                        custidname_pJP[m] = custid_pJP.ToString();
                        m += 1;
                        y = i + 1;
                        i++;
                    }

                    //填入excel欄位值
                    wsheet_JP.Cell(i + 5, 1).Value = row[0]; //客戶代號
                    wsheet_JP.Cell(i + 5, 2).Value = row[1]; //客戶簡稱
                    wsheet_JP.Cell(i + 5, 3).Value = row[2]; //銷貨日期
                    wsheet_JP.Cell(i + 5, 4).Style.NumberFormat.Format = "@";
                    wsheet_JP.Cell(i + 5, 4).Value = row[3]; //單別
                    wsheet_JP.Cell(i + 5, 5).Style.NumberFormat.Format = "@";
                    wsheet_JP.Cell(i + 5, 5).Value = row[4]; //單號
                    wsheet_JP.Cell(i + 5, 6).Value = row[5]; //批號
                    wsheet_JP.Cell(i + 5, 7).Value = row[6]; //品號
                    wsheet_JP.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                    wsheet_JP.Cell(i + 5, 8).Value = row[7]; //數量
                    wsheet_JP.Cell(i + 5, 9).Style.NumberFormat.Format = "#,##0";
                    wsheet_JP.Cell(i + 5, 9).Value = row[8]; //原幣未稅金額
                    wsheet_JP.Cell(i + 5, 10).Style.NumberFormat.Format = "#,##0";
                    wsheet_JP.Cell(i + 5, 10).Value = row[9]; //原幣稅額
                    wsheet_JP.Cell(i + 5, 11).Style.NumberFormat.Format = "#,##0";
                    wsheet_JP.Cell(i + 5, 11).Value = row[10]; //原幣合計金額
                    wsheet_JP.Cell(i + 5, 12).Value = row[11]; //幣別
                    wsheet_JP.Cell(i + 5, 13).Style.NumberFormat.Format = "#,##0.000";
                    wsheet_JP.Cell(i + 5, 13).Value = row[12]; //匯率
                    wsheet_JP.Cell(i + 5, 14).Style.NumberFormat.Format = "#,##0";
                    wsheet_JP.Cell(i + 5, 14).Value = row[13]; //本幣未稅金額
                    wsheet_JP.Cell(i + 5, 15).Style.NumberFormat.Format = "#,##0";
                    wsheet_JP.Cell(i + 5, 15).Value = row[14]; //本幣稅額
                    wsheet_JP.Cell(i + 5, 16).Style.NumberFormat.Format = "#,##0";
                    wsheet_JP.Cell(i + 5, 16).Value = row[15]; //本幣合計金額

                    cust_pJP = row[0].ToString().Trim();
                    custid_pJP = row[0].ToString().Trim();


                    if ((rows_count_pJP - 1) == dt_pJP.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        i++;
                        wsheet_JP.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_JP.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_JP.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_JP.Cell(i + 5, 1).Value = custid_pJP;
                        wsheet_JP.Cell(i + 5, 3).Value = "小計";
                        wsheet_JP.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_JP.Cell(i + 5, 8).FormulaA1 = "=sum(H" + (y + 5) + ":H" + (i + 4) + ")";
                        wsheet_JP.Range("N" + (i + 5) + ":P" + (i + 5)).Style.NumberFormat.Format = "#,##0";
                        wsheet_JP.Cell(i + 5, 14).FormulaA1 = "=sum(N" + (y + 5) + ":N" + (i + 4) + ")";
                        wsheet_JP.Cell(i + 5, 15).FormulaA1 = "=sum(O" + (y + 5) + ":O" + (i + 4) + ")";
                        wsheet_JP.Cell(i + 5, 16).FormulaA1 = "=sum(P" + (y + 5) + ":P" + (i + 4) + ")";

                        custidname_pJP[m] = custid_pJP.ToString();
                        y = i + 1;

                        wsheet_JP.Cell("B2").Value = txt_date_s.Text.ToString() + "~" + txt_date_e.Text.ToString();
                        wsheet_JP.Cell("B3").Value = createday;
                        wsheet_JP.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_JP.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_JP.Cell(i + 6, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_JP.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_JP.Cell(i + 6, 3).Value = "總計";
                        wsheet_JP.Cell(i + 6, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_JP.Cell(i + 6, 8).FormulaA1 = "=SUMIF(C:C,\"小計\",H:H)";
                        sumrow_pJP[0] = i + 6;
                        wsheet_JP.Range("N" + (i + 6) + ":P" + (i + 6)).Style.NumberFormat.Format = "#,##0";
                        wsheet_JP.Cell(i + 6, 14).FormulaA1 = "=SUMIF(C:C,\"小計\",N:N)";
                        wsheet_JP.Cell(i + 6, 15).FormulaA1 = "=SUMIF(C:C,\"小計\",O:O)";
                        wsheet_JP.Cell(i + 6, 16).FormulaA1 = "=SUMIF(C:C,\"小計\",P:P)";

                        wsheet_JP.Range("N" + (i + 10) + ":N" + (i + 22)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_JP.Cell(i + 10, 13).Value = "AU(AU-C5E)";
                        wsheet_JP.Cell(i + 10, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"AU\",C:C,\"小計\")";
                        sumrow_pJP[1] = i + 10;
                        wsheet_JP.Cell(i + 11, 13).Value = "AU-T(AU-C4A)";
                        wsheet_JP.Cell(i + 11, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"AU-T\",C:C,\"小計\")";

                        wsheet_JP.Cell(i + 12, 13).Value = "AU-TK(路竹)";
                        wsheet_JP.Cell(i + 12, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"AU-TK\",C:C,\"小計\")";

                        wsheet_JP.Cell(i + 13, 13).Value = "AU-TN(台南)";
                        wsheet_JP.Cell(i + 13, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"AU-TN\",C:C,\"小計\")";

                        wsheet_JP.Cell(i + 14, 13).Value = "TCE";
                        wsheet_JP.Cell(i + 14, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"TCE\",C:C,\"小計\")";

                        wsheet_JP.Cell(i + 15, 13).Value = "TCET-AU";
                        wsheet_JP.Cell(i + 15, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"TCET-AU\",C:C,\"小計\")";

                        for (int num = 0; num <= m; num++)
                        {
                            switch (custidname_pJP[num])
                            {
                                case "AU":
                                case "AU-T":
                                case "AU-TK":
                                case "AU-TN":
                                case "TCE":
                                case "TCET-AU":

                                case "CSOT":
                                case "WCSOT":
                                case "CPCF":
                                case "AU-L6K":
                                case "AU-TC":
                                case "AU-TY":
                                case "BOE":
                                case "CPD":
                                case "HKC":
                                case "HKC-H2":
                                case "HKC-H4":
                                case "HKC-H5":
                                case "CCPD":
                                case "CHOT":
                                case "HSD":
                                    break;
                                default:
                                    wsheet_JP.Cell(i + 16, 13).Value = custidname_pJP[num];
                                    wsheet_JP.Cell(i + 16, 14).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                                    wsheet_JP.Cell(i + 16, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"" + custidname_pJP[num] + "\",C:C,\"小計\")";
                                    i++;
                                    break;
                            }
                            p++;
                        }

                        wsheet_JP.Range("M" + (i + 16) + ":N" + (i + 19)).Style.Fill.BackgroundColor = XLColor.FromHtml("#F1DBDA");
                        wsheet_JP.Cell(i + 16, 13).Value = "折讓-2%";
                        wsheet_JP.Cell(i + 16, 14).FormulaA1 = "=-折讓_日本凸版!N" + num_dcJP[0];
                        sumrow_pJP[3] = i + 16;

                        wsheet_JP.Cell(i + 17, 13).Value = "折讓-1%";
                        wsheet_JP.Cell(i + 17, 14).FormulaA1 = "=-折讓_日本凸版!N" + num_dcJP[1];


                        wsheet_JP.Cell(i + 18, 13).Value = "匯款手續費-2%";
                        wsheet_JP.Cell(i + 18, 14).FormulaA1 = "=-手續費!H" + num_hJP[0];

                        wsheet_JP.Cell(i + 19, 13).Value = "匯款手續費-1%";
                        wsheet_JP.Cell(i + 19, 14).FormulaA1 = "=-手續費!H" + num_hJP[1];
                        sumrow_pJP[2] = i + 19;
                        sumrow_pJP[4] = i + 19;

                        wsheet_JP.Cell(i + 20, 13).Value = "合計";
                        wsheet_JP.Cell(i + 20, 14).Style.NumberFormat.Format = "#,##0";
                        wsheet_JP.Range("M" + (i + 20) + ":N" + (i + 20)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_JP.Cell(i + 20, 14).FormulaA1 = "=N" + sumrow_pJP[0] + "+SUM(N" + sumrow_pJP[1] + ":N" + sumrow_pJP[2] + ")";

                        wsheet_JP.Range("F" + (i + 24) + ":N" + (i + 27)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_JP.Range("F" + (i + 24) + ":N" + (i + 27)).Style.Font.FontSize = 14;
                        wsheet_JP.Range("K" + (i + 25) + ":N" + (i + 27)).Style.NumberFormat.Format = "#,##0";
                        wsheet_JP.Range("G" + (i + 25) + ":G" + (i + 26)).Style.Font.FontSize = 10;

                        wsheet_JP.Range("F" + (i + 24) + ":N" + (i + 27)).Style.Border.OutsideBorder = XLBorderStyleValues.Double;

                        wsheet_JP.Range("H" + (i + 24) + ":H" + (i + 26)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsheet_JP.Cell(i + 24, 11).Value = "銷貨";
                        wsheet_JP.Cell(i + 24, 12).Value = "權利金";
                        wsheet_JP.Cell(i + 24, 13).Value = "10%稅金";
                        wsheet_JP.Cell(i + 24, 14).Value = "淨額";

                        wsheet_JP.Cell(i + 25, 6).Value = "Royalty";
                        wsheet_JP.Cell(i + 25, 7).Value = "CSOT、WCSOT、CPCF、AU-L6K、AU-TC、AU-TY";
                        wsheet_JP.Cell(i + 25, 10).Value = "2%";
                        wsheet_JP.Cell(i + 25, 11).FormulaA1 = "=SUM(SUMIFS(N:N,A:A,{\"CSOT\",\"WCSOT\",\"CPCF\",\"AU-L6K\",\"AU-TC\",\"AU-TY\"},C:C,\"小計\"))" + str_enter +
                                                               "+SUM(SUMIFS(N" + sumrow_pJP[3] + ":N" + sumrow_pJP[4] + ",M" + sumrow_pJP[3] + ":M" + sumrow_pJP[4] + str_enter +
                                                               ",{\"折讓-2%\",\"匯款手續費-2%\"}))";

                        wsheet_JP.Cell(i + 25, 12).FormulaA1 = "= ROUND(K" + (i + 25) + "* 0.02, 0)";
                        wsheet_JP.Cell(i + 25, 13).FormulaA1 = "= ROUND(L" + (i + 25) + "* 0.1, 0)";
                        wsheet_JP.Cell(i + 25, 14).FormulaA1 = "= L" + (i + 25) + "-M" + (i + 25);

                        wsheet_JP.Cell(i + 26, 7).Value = "BOE、CPD、HKC、HKC-H2、HKC-H4、HKC-H5、CHOT、HSD、CCPD";
                        wsheet_JP.Cell(i + 26, 10).Value = "1%";
                        wsheet_JP.Cell(i + 26, 11).FormulaA1 = "=SUM(SUMIFS(N:N,A:A,{\"BOE\",\"CPD\",\"HKC\",\"HKC-H2\",\"HKC-H4\",\"HKC-H5\",\"CHOT\",\"HSD\",\"CCPD\"},C:C,\"小計\"))" + str_enter +
                                                               "+SUM(SUMIFS(N" + sumrow_pJP[3] + ":N" + sumrow_pJP[4] + ",M" + sumrow_pJP[3] + ":M" + sumrow_pJP[4] + str_enter +
                                                               ",{\"折讓-1%\",\"匯款手續費-1%\"}))";
                        wsheet_JP.Cell(i + 26, 12).FormulaA1 = "= ROUND(K" + (i + 26) + "* 0.01, 0)";
                        wsheet_JP.Cell(i + 26, 13).FormulaA1 = "= ROUND(L" + (i + 26) + "* 0.1, 0)";
                        wsheet_JP.Cell(i + 26, 14).FormulaA1 = "= L" + (i + 26) + "-M" + (i + 26);

                        wsheet_JP.Cell(i + 27, 8).Value = "2%+1%";
                        wsheet_JP.Cell(i + 27, 11).FormulaA1 = "=sum(K" + (i + 25) + ":K" + (i + 26) + ")";
                        wsheet_JP.Cell(i + 27, 12).FormulaA1 = "=sum(L" + (i + 25) + ":L" + (i + 26) + ")";
                        wsheet_JP.Cell(i + 27, 13).FormulaA1 = "=sum(M" + (i + 25) + ":M" + (i + 26) + ")";
                        wsheet_JP.Cell(i + 27, 14).FormulaA1 = "=sum(N" + (i + 25) + ":N" + (i + 26) + ")";
                    }
                    i++;
                }
                //worksheet.Columns().AdjustToContents();
                //worksheet2.Columns().AdjustToContents();

                wsheet_JP.Position = 1;
                wsheet_JP.Column("K").Width = 16;
                wsheet_JP.Column("L").Width = 16;
                wsheet_JP.Column("M").Width = 16;
                wsheet_JP.Column("N").Width = 16;

                save_as_JP = txt_path.Text.ToString().Trim() + @"\\銷貨明細(權利金)日本凸版" + txt_date_e.Text.ToString().Substring(0, 6) + "_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";

                wb_JP.SaveAs(save_as_JP);

            }
            //打开文件
            System.Diagnostics.Process.Start(save_as_JP);
            System.Diagnostics.Process.Start(save_as_MPT);
        }

        private void Btn_pre_Click(object sender, EventArgs e)
        {
            //dgv_p_all
            //     ClosedXMLExportExcel(DataGridView dgvDataInfo, string path, string fileNameWithExtension)
            //MyCode.ClosedXMLExportExcel(dgv_p_all, txt_path.Text.ToString().Trim(),"1.xlsx");
            using (XLWorkbook wb_pAll = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel))
                {
                    var ws = templateWB.Worksheet(5);

                    ws.CopyTo(wb_pAll, "權利金");
                }

                var wsheet_pAll = wb_pAll.Worksheet("權利金");

                //== 權利金報表 ================================================================
                //var worksheet = wb.Worksheets.Add("權利金_光阻");

                int rows_count_pAll = dt_pAll.Rows.Count;
                int i = 0; int y = 0;
                string cust_pAll = "";
                string sum_pAll = "";

                foreach (DataRow row in dt_pMPT.Rows)
                {
                    if (cust_pAll.ToString() != "" && row[0].ToString() != cust_pAll.ToString())
                    {
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_pAll.Cell(i + 5, 1).Value = cust_pAll;
                        wsheet_pAll.Cell(i + 5, 3).Value = "小計";
                        wsheet_pAll.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_pAll.Cell(i + 5, 8).FormulaA1 = "=sum(H" + (y + 5) + ":H" + (i + 4) + ")";
                        wsheet_pAll.Range("N" + (i + 5) + ":P" + (i + 5)).Style.NumberFormat.Format = "#,##0";
                        wsheet_pAll.Cell(i + 5, 14).FormulaA1 = "=sum(N" + (y + 5) + ":N" + (i + 4) + ")";
                        wsheet_pAll.Cell(i + 5, 15).FormulaA1 = "=sum(O" + (y + 5) + ":O" + (i + 4) + ")";
                        wsheet_pAll.Cell(i + 5, 16).FormulaA1 = "=sum(P" + (y + 5) + ":P" + (i + 4) + ")";

                        y = i + 1;
                        i++;

                    }

                    //填入excel欄位值
                    wsheet_pAll.Cell(i + 5, 1).Value = row[0]; //客戶代號
                    wsheet_pAll.Cell(i + 5, 2).Value = row[1]; //客戶簡稱
                    wsheet_pAll.Cell(i + 5, 3).Value = row[2]; //銷貨日期
                    wsheet_pAll.Cell(i + 5, 4).Style.NumberFormat.Format = "@";
                    wsheet_pAll.Cell(i + 5, 4).Value = row[3]; //單別
                    wsheet_pAll.Cell(i + 5, 5).Style.NumberFormat.Format = "@";
                    wsheet_pAll.Cell(i + 5, 5).Value = row[4]; //單號
                    wsheet_pAll.Cell(i + 5, 6).Value = row[5]; //批號
                    wsheet_pAll.Cell(i + 5, 7).Value = row[6]; //品號

                    if (row[6].ToString().Substring(0, 3) == "MPT")
                    {
                        sum_pAll += "N" + (i + 5) + "+";
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDD7EE");
                    }

                    wsheet_pAll.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                    wsheet_pAll.Cell(i + 5, 8).Value = row[7]; //數量
                    wsheet_pAll.Cell(i + 5, 9).Style.NumberFormat.Format = "#,##0";
                    wsheet_pAll.Cell(i + 5, 9).Value = row[8]; //原幣未稅金額
                    wsheet_pAll.Cell(i + 5, 10).Style.NumberFormat.Format = "#,##0";
                    wsheet_pAll.Cell(i + 5, 10).Value = row[9]; //原幣稅額
                    wsheet_pAll.Cell(i + 5, 11).Style.NumberFormat.Format = "#,##0";
                    wsheet_pAll.Cell(i + 5, 11).Value = row[10]; //原幣合計金額
                    wsheet_pAll.Cell(i + 5, 12).Value = row[11]; //幣別
                    wsheet_pAll.Cell(i + 5, 13).Style.NumberFormat.Format = "#,##0.000";
                    wsheet_pAll.Cell(i + 5, 13).Value = row[12]; //匯率
                    wsheet_pAll.Cell(i + 5, 14).Style.NumberFormat.Format = "#,##0";
                    wsheet_pAll.Cell(i + 5, 14).Value = row[13]; //本幣未稅金額
                    wsheet_pAll.Cell(i + 5, 15).Style.NumberFormat.Format = "#,##0";
                    wsheet_pAll.Cell(i + 5, 15).Value = row[14]; //本幣稅額
                    wsheet_pAll.Cell(i + 5, 16).Style.NumberFormat.Format = "#,##0";
                    wsheet_pAll.Cell(i + 5, 16).Value = row[15]; //本幣合計金額

                    cust_pAll = row[0].ToString().Trim();

                    if ((rows_count_pAll - 1) == dt_pMPT.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        i++;
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_pAll.Cell(i + 5, 1).Value = cust_pAll;
                        wsheet_pAll.Cell(i + 5, 3).Value = "小計";
                        wsheet_pAll.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_pAll.Cell(i + 5, 8).FormulaA1 = "=sum(H" + (y + 5) + ":H" + (i + 4) + ")";
                        wsheet_pAll.Range("N" + (i + 5) + ":P" + (i + 5)).Style.NumberFormat.Format = "#,##0";
                        wsheet_pAll.Cell(i + 5, 14).FormulaA1 = "=sum(N" + (y + 5) + ":N" + (i + 4) + ")";
                        wsheet_pAll.Cell(i + 5, 15).FormulaA1 = "=sum(O" + (y + 5) + ":O" + (i + 4) + ")";
                        wsheet_pAll.Cell(i + 5, 16).FormulaA1 = "=sum(P" + (y + 5) + ":P" + (i + 4) + ")";

                        y = i + 1;

                        wsheet_pAll.Cell("B2").Value = txt_date_s.Text.ToString() + "~" + txt_date_e.Text.ToString();
                        wsheet_pAll.Cell("B3").Value = createday;
                        wsheet_pAll.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_pAll.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_pAll.Cell(i + 6, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_pAll.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_pAll.Cell(i + 6, 3).Value = "總計";
                        wsheet_pAll.Cell(i + 6, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_pAll.Cell(i + 6, 8).FormulaA1 = "=SUMIF(C:C,\"小計\",H:H)";
                        wsheet_pAll.Range("N" + (i + 6) + ":P" + (i + 6)).Style.NumberFormat.Format = "#,##0";
                        wsheet_pAll.Cell(i + 6, 14).FormulaA1 = "=SUMIF(C:C,\"小計\",N:N)";
                        wsheet_pAll.Cell(i + 6, 15).FormulaA1 = "=SUMIF(C:C,\"小計\",O:O)";
                        wsheet_pAll.Cell(i + 6, 16).FormulaA1 = "=SUMIF(C:C,\"小計\",P:P)";
                    }
                    i++;
                }
                //wsheet_pAll.Columns().AdjustToContents();
                //worksheet2.Columns().AdjustToContents();

                wsheet_pAll.Position = 1;
                //wsheet_pAll.Column("K").Width = 16;
                //wsheet_pAll.Column("L").Width = 16;
                //wsheet_pAll.Column("M").Width = 16;
                //wsheet_pAll.Column("N").Width = 16;

                save_as_All = txt_path.Text.ToString().Trim() + @"\\權利金" + txt_date_e.Text.ToString().Substring(0, 6) + "_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";

                wb_pAll.SaveAs(save_as_All);

                //打开文件
                System.Diagnostics.Process.Start(save_as_All);
            }

        }

        private void Btn_dc_Click(object sender, EventArgs e)
        {
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
                        wsheet_dcAll.Cell("B5").Value = createday;
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

                save_as_dcAll = txt_path.Text.ToString().Trim() + @"\\折讓" + txt_date_e.Text.ToString().Substring(0, 6) + "_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";

                wb_dcAll.SaveAs(save_as_dcAll);

                //打开文件
                System.Diagnostics.Process.Start(save_as_dcAll);
            }
        }

        private void btn_fileopen_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process prc = new System.Diagnostics.Process();
            prc.StartInfo.FileName = txt_path.Text.ToString();
            prc.Start();
        }

        private void btn_fileopen2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process prc2 = new System.Diagnostics.Process();
            prc2.StartInfo.FileName = txt_path2.Text.ToString();
            prc2.Start();
        }

        private void btn_search2_Click(object sender, EventArgs e)
        {
            Btn_acc2.Enabled = false;
            Btn_pre2.Enabled = false;
            Btn_dc2.Enabled = false;

            date_s2 = DateTime.ParseExact(txt_date_s2.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e2 = DateTime.ParseExact(txt_date_e2.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            if (date_s2 > date_e2)
            {
                MessageBox.Show("請修改日期區間", "日期格式錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //月份條件
            /*①CHOT    3月15日、6月15日、9月15日、12月15日以後的出貨，算在下一個月的營業額
              15 - 31日列隔月營業額  ,4月.7月.10月.1月加回上個月扣除的營業額,重新計算
            ②CPCF.CSOT.WCSOT  3月21日、6月21日、9月21日、12月21日以後的出貨，算在下一個月的營業額
              21 - 31日列隔月營業額  ,4月.7月.10月.1月加回上個月扣除的營業額,重新計算*/
              str_date_e2_m = txt_date_e2.Text.Trim().Substring(4, 2);
              date_e2 = DateTime.ParseExact(txt_date_e2.Text.Trim(), "yyyyMMdd", null);

            switch (str_date_e2_m)
            {
  /*              case "03":
                case "06":
                case "09":
                case "12":
                    str_date_s2_15 = txt_date_e2.Text.Trim().Substring(0, 6) + "15";
                    str_date_e2_15 = "";

                    str_date_s2_21 = "";
                    str_date_e2_21 = "";
                    break;
*/
                case "01":
                case "04":
                case "07":
                case "10":
                    str_date_s2 = txt_date_s2.Text.Trim();
                    str_date_e2 = txt_date_e2.Text.Trim();

                    str_date_s2_15 = DateTime.Parse(date_e2.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMM15");
                    str_date_e2_15 = txt_date_e2.Text.Trim();

                    str_date_s2_21 = DateTime.Parse(date_e2.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMM21");
                    str_date_e2_21 = txt_date_e2.Text.Trim();
                    break;
                default:
                    str_date_s2 = txt_date_s2.Text.Trim();
                    str_date_e2 = txt_date_e2.Text.Trim();

                    str_date_s2_15 = txt_date_s2.Text.Trim();
                    str_date_e2_15 = txt_date_e2.Text.Trim();

                    str_date_s2_21 = txt_date_s2.Text.Trim();
                    str_date_e2_21 = txt_date_e2.Text.Trim();
                    break;
            }

            if (dgv_p_all.DataSource != null) 
            {
                dt_pAll.Clear();
                dt_dcAll.Clear();
                dt_pMPT.Clear();
                dt_dcMPT.Clear();
                dt_pJP.Clear();
                dt_dcJP.Clear();
                dt_hJP.Clear();

                dgv_p_all.DataSource = null;
                dgv_dc_all.DataSource = null;
                dgv_p_MPT.DataSource = null;
                dgv_dc_MPT.DataSource = null;
                dgv_p_JP.DataSource = null;
                dgv_dc_JP.DataSource = null;
                dgv_hJP.DataSource = null;
            }
            //dt_pAll.Clear();
            //dt_dcAll.Clear();
            //dt_pMPT.Clear();
            //dt_dcMPT.Clear();
            //dt_pJP.Clear();
            //dt_dcJP.Clear();
            //dt_hJP.Clear();

            //dgv_p_all.DataSource = null;
            //dgv_dc_all.DataSource = null;
            //dgv_p_MPT.DataSource = null;
            //dgv_dc_MPT.DataSource = null;
            //dgv_p_JP.DataSource = null;
            //dgv_dc_JP.DataSource = null;
            //dgv_hJP.DataSource = null;

            // 權利金-光阻 = 銷貨日報表全部
            // COPTH.TH020 = 'Y'    確認碼為"Y"
            // COPTG.TG010 = '002'  廠別為"台南"
            // COPTH.TH001 in ('230', '2302', '230T', '234T', '235T')  篩選單別【'230', '2302', '230T', '234T', '235T'】
            // COPTH.TH017 like 'T%'    抓取批號為【T開頭】
            string sql_str_pMPT2 = String.Format(
                @"SELECT COPTG.TG004 as 客戶代號,COPMA.MA002 as 客戶簡稱,COPTG.TG003 as 銷貨日期,COPTH.TH001 as 單別,
                COPTH.TH002 as 單身單號,COPTH.TH017 as 批號,COPTH.TH004 as 品號,COPTH.TH008 as 數量,
                COPTH.TH035 as 原幣未稅金額,COPTH.TH036 as 原幣稅額 ,COPTH.TH035 + COPTH.TH036 as 原幣合計金額,
                COPTG.TG011 as 幣別,COPTG.TG012 as 匯率,
                COPTH.TH037 as 本幣未稅金額,COPTH.TH038 as 本幣稅額,COPTH.TH037 + COPTH.TH038 as 本幣合計金額
                FROM COPTH as COPTH
                Inner JOIN COPTG as COPTG On COPTH.TH001 = COPTG.TG001 and COPTH.TH002 = COPTG.TG002
                Inner JOIN COPMA as COPMA On COPTG.TG004 = COPMA.MA001
                where COPTG.TG003 >= '{0}' and COPTG.TG003 <= '{1}' and COPTH.TH020 = 'Y' and COPTG.TG010 = '002'
                and COPTH.TH001 in ({6})
                and COPTH.TH017 like 'T%' and COPTG.TG004 not in ({9},{10})
                union all
                SELECT COPTG.TG004 as 客戶代號,COPMA.MA002 as 客戶簡稱,COPTG.TG003 as 銷貨日期,COPTH.TH001 as 單別,
                                COPTH.TH002 as 單身單號,COPTH.TH017 as 批號,COPTH.TH004 as 品號,COPTH.TH008 as 數量,
                                COPTH.TH035 as 原幣未稅金額,COPTH.TH036 as 原幣稅額 ,COPTH.TH035 + COPTH.TH036 as 原幣合計金額,
                                COPTG.TG011 as 幣別,COPTG.TG012 as 匯率,
                                COPTH.TH037 as 本幣未稅金額,COPTH.TH038 as 本幣稅額,COPTH.TH037 + COPTH.TH038 as 本幣合計金額
                                FROM COPTH as COPTH
                                Inner JOIN COPTG as COPTG On COPTH.TH001 = COPTG.TG001 and COPTH.TH002 = COPTG.TG002
                                Inner JOIN COPMA as COPMA On COPTG.TG004 = COPMA.MA001
                                where COPTG.TG003 >= '{2}' and COPTG.TG003 <= '{3}' and COPTH.TH020 = 'Y' and COPTG.TG010 = '002'
                                and COPTH.TH001 in ({7})
                                and COPTH.TH017 like 'T%' and COPTG.TG004 in ({11})
                union all
                SELECT COPTG.TG004 as 客戶代號,COPMA.MA002 as 客戶簡稱,COPTG.TG003 as 銷貨日期,COPTH.TH001 as 單別,
                                COPTH.TH002 as 單身單號,COPTH.TH017 as 批號,COPTH.TH004 as 品號,COPTH.TH008 as 數量,
                                COPTH.TH035 as 原幣未稅金額,COPTH.TH036 as 原幣稅額 ,COPTH.TH035 + COPTH.TH036 as 原幣合計金額,
                                COPTG.TG011 as 幣別,COPTG.TG012 as 匯率,
                                COPTH.TH037 as 本幣未稅金額,COPTH.TH038 as 本幣稅額,COPTH.TH037 + COPTH.TH038 as 本幣合計金額
                                FROM COPTH as COPTH
                                Inner JOIN COPTG as COPTG On COPTH.TH001 = COPTG.TG001 and COPTH.TH002 = COPTG.TG002
                                Inner JOIN COPMA as COPMA On COPTG.TG004 = COPMA.MA001
                                where COPTG.TG003 >= '{4}' and COPTG.TG003 <= '{5}' and COPTH.TH020 = 'Y' and COPTG.TG010 = '002'
                                and COPTH.TH001 in ({8})
                                and COPTH.TH017 like 'T%' and COPTG.TG004 in ({12})
                ORDER BY 客戶代號 asc, 銷貨日期 asc, 單身單號 asc, 品號 asc"
                , str_date_s2, str_date_e2, str_date_s2_15, str_date_e2_15, str_date_s2_21, str_date_e2_21
                ,copth_th001, copth_th001, copth_th001, coptg_tg004_15, coptg_tg004_21, coptg_tg004_15, coptg_tg004_21);

            MyCode.Sql_dgv(sql_str_pMPT2, dt_pMPT, dgv_p_MPT);
            MyCode.Sql_dgv(sql_str_pMPT2, dt_pAll, dgv_p_all);

            // 權利金-光阻折讓
            //240 	折讓單-製品
            //240T 折讓單-製品 - 關係人
            //242     折讓單 - 商品
            //242T 折讓單-商品 - 關係人
            //批號為 T開頭，單別為[240.240T.241.241T]
            string sql_str_dcMPT2 = String.Format(
                @"SELECT COPTI.TI004 as 客戶,COPTG.TG003 as 銷貨日期,
                COPTI.TI001 as 折讓單別,COPTI.TI002 as 折讓單號,COPTC.TC012 as 客戶單號,
                COPTJ.TJ004 as 品號,COPTJ.TJ014 as 批號,COPTH.TH012 as 銷貨單價,COPTJ.TJ011 as 新單價,
                (COPTH.TH012 - COPTJ.TJ011) as 折讓差,COPTH.TH008 as 銷貨數量,
                COPTJ.TJ031 as 折讓金額,COPTJ.TJ032 as 折讓稅額,COPTJ.TJ033 as 台幣金額,COPTJ.TJ034 as 台幣稅額,COPTI.TI009 as 匯率,
                COPTJ.TJ033 + COPTJ.TJ034 as 台幣合計,COPTI.TI014 as 發票號碼,COPTI.TI006 as 廠別,COPTI.TI003 as 銷退日期
                FROM COPTJ as COPTJ
                Left JOIN COPTH as COPTH On COPTJ.TJ015 = COPTH.TH001 and COPTJ.TJ016 = COPTH.TH002 and COPTJ.TJ017 = COPTH.TH003
                Left JOIN COPTI as COPTI On COPTJ.TJ001 = COPTI.TI001 and COPTJ.TJ002 = COPTI.TI002
                Left JOIN COPTC as COPTC On COPTH.TH014 = COPTC.TC001 AND COPTH.TH015 = COPTC.TC002
                Left JOIN COPTG as COPTG On COPTH.TH001 = COPTG.TG001 AND COPTH.TH002 = COPTG.TG002
                where COPTI.TI003 >= '{0}' and COPTI.TI003 <= '{1}'
                and COPTJ.TJ014 like 'T%' and COPTI.TI001 in ({2})
                ORDER BY COPTI.TI003 asc, COPTI.TI004 asc, COPTG.TG003 asc"
                , str_date_s2, str_date_e2,dcMPT_copti_ti001);

            MyCode.Sql_dgv(sql_str_dcMPT2, dt_dcMPT, dgv_dc_MPT);
            MyCode.Sql_dgv(sql_str_dcMPT2, dt_dcAll, dgv_dc_all);


            // 權利金-日本凸版 
            // 同銷貨日報表條件，新增排除 品名為光阻 MPT 品項

            string sql_str_pJP2 = String.Format(
                @"SELECT COPTG.TG004 as 客戶代號,COPMA.MA002 as 客戶簡稱,COPTG.TG003 as 銷貨日期,COPTH.TH001 as 單別,
                COPTH.TH002 as 單身單號,COPTH.TH017 as 批號,COPTH.TH004 as 品號,COPTH.TH008 as 數量,
                COPTH.TH035 as 原幣未稅金額,COPTH.TH036 as 原幣稅額   ,COPTH.TH035 + COPTH.TH036 as 原幣合計金額,
                COPTG.TG011 as 幣別,COPTG.TG012 as 匯率,
                COPTH.TH037 as 本幣未稅金額,COPTH.TH038 as 本幣稅額,COPTH.TH037 + COPTH.TH038 as 本幣合計金額
                FROM COPTH as COPTH
                Inner JOIN COPTG as COPTG On COPTH.TH001 = COPTG.TG001 and COPTH.TH002 = COPTG.TG002
                Inner JOIN COPMA as COPMA On COPTG.TG004 = COPMA.MA001
                where COPTG.TG003 >= '{0}' and COPTG.TG003 <= '{1}' and COPTH.TH020 = 'Y' and COPTG.TG010 = '002'
                and COPTH.TH001 in ({6})
                and COPTH.TH017 like 'T%' and COPTH.TH004 not like 'MPT%' and COPTG.TG004 not in ({9},{10})
                union all
                SELECT COPTG.TG004 as 客戶代號,COPMA.MA002 as 客戶簡稱,COPTG.TG003 as 銷貨日期,COPTH.TH001 as 單別,
                                COPTH.TH002 as 單身單號,COPTH.TH017 as 批號,COPTH.TH004 as 品號,COPTH.TH008 as 數量,
                                COPTH.TH035 as 原幣未稅金額,COPTH.TH036 as 原幣稅額   ,COPTH.TH035 + COPTH.TH036 as 原幣合計金額,
                                COPTG.TG011 as 幣別,COPTG.TG012 as 匯率,
                                COPTH.TH037 as 本幣未稅金額,COPTH.TH038 as 本幣稅額,COPTH.TH037 + COPTH.TH038 as 本幣合計金額
                                FROM COPTH as COPTH
                                Inner JOIN COPTG as COPTG On COPTH.TH001 = COPTG.TG001 and COPTH.TH002 = COPTG.TG002
                                Inner JOIN COPMA as COPMA On COPTG.TG004 = COPMA.MA001
                                where COPTG.TG003 >= '{2}' and COPTG.TG003 <= '{3}' and COPTH.TH020 = 'Y' and COPTG.TG010 = '002'
                                and COPTH.TH001 in ({7})
                                and COPTH.TH017 like 'T%' and COPTH.TH004 not like 'MPT%' and COPTG.TG004 in ({11})
                union all
                SELECT COPTG.TG004 as 客戶代號,COPMA.MA002 as 客戶簡稱,COPTG.TG003 as 銷貨日期,COPTH.TH001 as 單別,
                                COPTH.TH002 as 單身單號,COPTH.TH017 as 批號,COPTH.TH004 as 品號,COPTH.TH008 as 數量,
                                COPTH.TH035 as 原幣未稅金額,COPTH.TH036 as 原幣稅額   ,COPTH.TH035 + COPTH.TH036 as 原幣合計金額,
                                COPTG.TG011 as 幣別,COPTG.TG012 as 匯率,
                                COPTH.TH037 as 本幣未稅金額,COPTH.TH038 as 本幣稅額,COPTH.TH037 + COPTH.TH038 as 本幣合計金額
                                FROM COPTH as COPTH
                                Inner JOIN COPTG as COPTG On COPTH.TH001 = COPTG.TG001 and COPTH.TH002 = COPTG.TG002
                                Inner JOIN COPMA as COPMA On COPTG.TG004 = COPMA.MA001
                                where COPTG.TG003 >= '{4}' and COPTG.TG003 <= '{5}' and COPTH.TH020 = 'Y' and COPTG.TG010 = '002'
                                and COPTH.TH001 in ({8})
                                and COPTH.TH017 like 'T%' and COPTH.TH004 not like 'MPT%' and COPTG.TG004 in ({12})
                ORDER BY 客戶代號 asc, 銷貨日期 asc, 單身單號 asc, 品號 asc"
                , str_date_s2, str_date_e2, str_date_s2_15, str_date_e2_15, str_date_s2_21, str_date_e2_21
                , copth_th001, copth_th001, copth_th001, coptg_tg004_15, coptg_tg004_21, coptg_tg004_15, coptg_tg004_21);


            MyCode.Sql_dgv(sql_str_pJP2, dt_pJP, dgv_p_JP);

            // 權利金-日本凸版折讓
            //批號 T開頭，品號不等於'MPT'，單別為[240.241]
            string sql_str_dcJP2 = String.Format(
                @"SELECT COPTI.TI004 as 客戶,COPTG.TG003 as 銷貨日期,
               COPTI.TI001 as 折讓單別,COPTI.TI002 as 折讓單號,COPTC.TC012 as 客戶單號,
               COPTJ.TJ004 as 品號,COPTJ.TJ014 as 批號,COPTH.TH012 as 銷貨單價,COPTJ.TJ011 as 新單價,
               (COPTH.TH012 - COPTJ.TJ011) as 折讓差,COPTH.TH008 as 銷貨數量,
               COPTJ.TJ031 as 折讓金額,COPTJ.TJ032 as 折讓稅額,COPTJ.TJ033 as 台幣金額,COPTJ.TJ034 as 台幣稅額,COPTI.TI009 as 匯率,
               COPTJ.TJ033 + COPTJ.TJ034 as 台幣合計,COPTI.TI014 as 發票號碼,COPTI.TI006 as 廠別,COPTI.TI003 as 銷退日期
               FROM COPTJ as COPTJ
               Left JOIN COPTH as COPTH On COPTJ.TJ015 = COPTH.TH001 and COPTJ.TJ016 = COPTH.TH002 and COPTJ.TJ017 = COPTH.TH003
               Left JOIN COPTI as COPTI On COPTJ.TJ001 = COPTI.TI001 and COPTJ.TJ002 = COPTI.TI002
               Left JOIN COPTC as COPTC On COPTH.TH014 = COPTC.TC001 AND COPTH.TH015 = COPTC.TC002
               Left JOIN COPTG as COPTG On COPTH.TH001 = COPTG.TG001 AND COPTH.TH002 = COPTG.TG002
               where COPTI.TI003 >= '{0}' and COPTI.TI003 <= '{1}' and COPTJ.TJ014 like 'T%' and COPTH.TH004 not like 'MPT%'
               and COPTI.TI001 in ({2})
               ORDER BY COPTI.TI003 asc, COPTI.TI004 asc, COPTG.TG003 asc"
               , str_date_s2, str_date_e2,dcJP_copti_ti001);

            MyCode.Sql_dgv(sql_str_dcJP2, dt_dcJP, dgv_dc_JP);

            //手續費
            //依摘要 篩選
            string sql_str_hJP2 = String.Format(
                @"select * from 
                    (select SUBSTRING( ML009 ,1, CHARINDEX (' ', ML009) -1) as 客戶名稱, 
                    ML001 as 科目編號 ,
                    (select MA003 from ACTMA where MA001 = ACTML.ML001) as 科目名稱,
                    ML002 as 傳票日期 ,
                    ML003+'-'+ML004+' -'+ML005 as 傳票編號,
                    ML009 as 摘要 ,
                    (case (SUBSTRING( ML009 ,1, CHARINDEX (' ', ML009) -1)) 
	                    when 'CSOT' then '2%'
	                    when 'WCSOT' then '2%'
	                    when 'CPCF' then '2%'
	                    when 'AU-L6K' then '2%'
	                    when 'AU-TC' then '2%'
	                    when 'AU-TY' then '2%'
	                    when 'BOE' then '1%'
	                    when 'CPD' then '1%'
	                    when 'HKC' then '1%'
	                    when 'HKC-H2' then '1%'
                        when 'HKC-H4' then '1%'
                        when 'HKC-H5' then '1%'
                        when 'CCPD' then '1%'
	                    when 'CHOT' then '1%'
                        when 'HSD' then '1%' end) as 手續費率,
                    (case ML007 when '1' then ML008 else 0 end) as 借方金額,
                    (case ML007 when '-1' then ML008 else 0 end)  as 貸方金額 ,
                    (case ML007 when '1' then '借餘' when '-1' then '貸餘' end) as 借貸 
                    from ACTML
                    where ML006 = '623202' and ML002 >='{0}' and ML002 <= '{1}' and ML009 like '%帳款手續費%') 手續費
                    where 手續費率 is not null
                    order by 手續費率 desc,客戶名稱,傳票日期"
                , str_date_s2, str_date_e2);

            MyCode.Sql_dgv(sql_str_hJP2, dt_hJP, dgv_hJP);

            tabControl1.SelectedIndex = 0;
            Btn_acc2.Enabled = true;
            Btn_pre2.Enabled = true;
            Btn_dc2.Enabled = true;
        }

        private void Btn_acc2_Click(object sender, EventArgs e)
        {
            if (dgv_p_MPT.DataSource is null)
            {
                MessageBox.Show("請先【查詢】", "無法轉出Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //保存成文件 權利金_光阻
            using (XLWorkbook wb_MPT = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel))
                {
                    var ws = templateWB.Worksheet(2);
                    var ws2 = templateWB.Worksheet(3);

                    ws.CopyTo(wb_MPT, "權利金_光阻");
                    ws2.CopyTo(wb_MPT, "折讓_光阻");
                }

                var wsheet_MPT = wb_MPT.Worksheet("權利金_光阻");
                var wsheet_dcMPT = wb_MPT.Worksheet("折讓_光阻");

                //== 折讓_光阻 ================================================================
                int rows_count_dcMPT = dt_dcMPT.Rows.Count;
                int j = 0; int x = 0;
                string cust_dcMPT = "";
                string sum_dcMPT = "";

                foreach (DataRow row in dt_dcMPT.Rows)
                {
                    if (cust_dcMPT.ToString() != "" && row[0].ToString() != cust_dcMPT.ToString())
                    {
                        wsheet_dcMPT.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_dcMPT.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_dcMPT.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_dcMPT.Cell(j + 7, 1).Value = cust_dcMPT;
                        wsheet_dcMPT.Cell(j + 7, 2).Value = "小計";
                        wsheet_dcMPT.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 7, 11).FormulaA1 = "=sum(K" + (x + 7) + ":K" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 12).FormulaA1 = "=sum(L" + (x + 7) + ":L" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 13).FormulaA1 = "=sum(M" + (x + 7) + ":M" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 14).FormulaA1 = "=sum(N" + (x + 7) + ":N" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 15).FormulaA1 = "=sum(O" + (x + 7) + ":O" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 7, 17).FormulaA1 = "=sum(Q" + (x + 7) + ":Q" + (j + 6) + ")";

                        sum_dcMPT += "N" + (j + 7) + "+";

                        x = j + 1;
                        j++;
                        //        cust = "1";
                    }

                    if (row[0].ToString() != cust_dcMPT.ToString())
                    {
                        wsheet_dcMPT.Cell(j + 7, 1).Value = row[0]; //客戶
                    }

                    wsheet_dcMPT.Cell(j + 7, 2).Value = row[1]; //銷貨日期
                    wsheet_dcMPT.Cell(j + 7, 3).Style.NumberFormat.Format = "@";
                    wsheet_dcMPT.Cell(j + 7, 3).Value = row[2]; //折讓單別
                    wsheet_dcMPT.Cell(j + 7, 4).Style.NumberFormat.Format = "@";
                    wsheet_dcMPT.Cell(j + 7, 4).Value = row[3]; //折讓單號
                    wsheet_dcMPT.Cell(j + 7, 5).Style.NumberFormat.Format = "@";
                    wsheet_dcMPT.Cell(j + 7, 5).Value = row[4]; //客戶單號
                    wsheet_dcMPT.Cell(j + 7, 6).Value = row[5]; //品號
                    wsheet_dcMPT.Cell(j + 7, 7).Value = row[6]; //批號
                    wsheet_dcMPT.Cell(j + 7, 8).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 8).Value = row[7]; //銷貨單價
                    wsheet_dcMPT.Cell(j + 7, 9).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 9).Value = row[8]; //新單價
                    wsheet_dcMPT.Cell(j + 7, 10).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 10).Value = row[9]; //折讓差

                    wsheet_dcMPT.Cell(j + 7, 11).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 11).Value = row[10]; //銷貨數量
                    wsheet_dcMPT.Cell(j + 7, 12).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 12).Value = row[11]; //折讓金額
                    wsheet_dcMPT.Cell(j + 7, 13).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 13).Value = row[12]; //折讓稅額
                    wsheet_dcMPT.Cell(j + 7, 14).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 14).Value = row[13]; //台幣金額
                    wsheet_dcMPT.Cell(j + 7, 15).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 15).Value = row[14]; //台幣稅額
                    wsheet_dcMPT.Cell(j + 7, 16).Style.NumberFormat.Format = "#,##0.0000";
                    wsheet_dcMPT.Cell(j + 7, 16).Value = row[15]; //匯率
                    wsheet_dcMPT.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcMPT.Cell(j + 7, 17).Value = row[16]; //台幣合計
                    wsheet_dcMPT.Cell(j + 7, 18).Value = row[17]; //發票號碼
                    wsheet_dcMPT.Cell(j + 7, 19).Value = row[18]; //廠別

                    wsheet_dcMPT.Cell(j + 7, 20).Value = row[19]; //銷退日

                    cust_dcMPT = row[0].ToString().Trim();

                    if ((rows_count_dcMPT - 1) == dt_dcMPT.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        j++;
                        wsheet_dcMPT.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_dcMPT.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_dcMPT.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_dcMPT.Cell(j + 7, 1).Value = cust_dcMPT;
                        wsheet_dcMPT.Cell(j + 7, 2).Value = "小計";
                        wsheet_dcMPT.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 7, 11).FormulaA1 = "=sum(K" + (x + 7) + ":K" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 12).FormulaA1 = "=sum(L" + (x + 7) + ":L" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 13).FormulaA1 = "=sum(M" + (x + 7) + ":M" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 14).FormulaA1 = "=sum(N" + (x + 7) + ":N" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 15).FormulaA1 = "=sum(O" + (x + 7) + ":O" + (j + 6) + ")";
                        wsheet_dcMPT.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 7, 17).FormulaA1 = "=sum(Q" + (x + 7) + ":Q" + (j + 6) + ")";

                        sum_dcMPT += "N" + (j + 7);
                        x = j + 1;

                        j++;

                        wsheet_dcMPT.Cell("B4").Value = txt_date_s2.Text.ToString() + " ~ " + txt_date_e2.Text.ToString();
                        wsheet_dcMPT.Cell("B5").Value = createday;
                        wsheet_dcMPT.Cell(j + 7, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_dcMPT.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_dcMPT.Cell(j + 7, 3).Value = "總計";
                        wsheet_dcMPT.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 7, 11).FormulaA1 = "=SUMIF(B:B,\"小計\",K:K)";
                        wsheet_dcMPT.Cell(j + 7, 12).FormulaA1 = "=SUMIF(B:B,\"小計\",L:L)";
                        wsheet_dcMPT.Cell(j + 7, 13).FormulaA1 = "=SUMIF(B:B,\"小計\",M:M)";
                        wsheet_dcMPT.Cell(j + 7, 14).FormulaA1 = "=SUMIF(B:B,\"小計\",N:N)";
                        wsheet_dcMPT.Cell(j + 7, 15).FormulaA1 = "=SUMIF(B:B,\"小計\",O:O)";
                        wsheet_dcMPT.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 7, 17).FormulaA1 = "=SUMIF(B:B,\"小計\",Q:Q)";

                        j = j + 2;
                        wsheet_dcMPT.Cell(j + 9, 13).Value = "UV製品-折讓";
                        wsheet_dcMPT.Cell(j + 10, 13).Value = " 製品-折讓";
                        wsheet_dcMPT.Cell(j + 10, 14).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 10, 14).FormulaA1 = "=SUMIF(B:B,\"小計\",N:N)";
                        wsheet_dcMPT.Cell(j + 11, 13).Value = " 製品-退回";
                        wsheet_dcMPT.Cell(j + 12, 13).Value = " 商品-折讓";
                        wsheet_dcMPT.Cell(j + 13, 13).Value = " 商品-退回";
                        wsheet_dcMPT.Cell(j + 14, 13).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_dcMPT.Range("M" + (j + 14) + ":N" + (j + 14)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_dcMPT.Cell(j + 14, 13).Value = "合計";
                        wsheet_dcMPT.Cell(j + 14, 14).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcMPT.Cell(j + 14, 14).FormulaA1 = "=sum(N" + (j + 10) + ":N" + (j + 13) + ")";
                    }
                    j++;
                }

                //== 權利金-光阻 ================================================================
                //var worksheet = wb.Worksheets.Add("權利金_光阻");

                int rows_count_pMPT = dt_pMPT.Rows.Count;
                int i = 0; int y = 0;
                string cust_pMPT = "";
                string sum_pMPT = "";

                foreach (DataRow row in dt_pMPT.Rows)
                {
                    if (cust_pMPT.ToString() != "" && row[0].ToString() != cust_pMPT.ToString())
                    {
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_MPT.Cell(i + 5, 1).Value = cust_pMPT;
                        wsheet_MPT.Cell(i + 5, 3).Value = "小計";
                        wsheet_MPT.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_MPT.Cell(i + 5, 8).FormulaA1 = "=sum(H" + (y + 5) + ":H" + (i + 4) + ")";
                        wsheet_MPT.Range("N" + (i + 5) + ":P" + (i + 5)).Style.NumberFormat.Format = "#,##0";
                        wsheet_MPT.Cell(i + 5, 14).FormulaA1 = "=sum(N" + (y + 5) + ":N" + (i + 4) + ")";
                        wsheet_MPT.Cell(i + 5, 15).FormulaA1 = "=sum(O" + (y + 5) + ":O" + (i + 4) + ")";
                        wsheet_MPT.Cell(i + 5, 16).FormulaA1 = "=sum(P" + (y + 5) + ":P" + (i + 4) + ")";

                        y = i + 1;
                        i++;

                    }

                    //填入excel欄位值
                    wsheet_MPT.Cell(i + 5, 1).Value = row[0]; //客戶代號
                    wsheet_MPT.Cell(i + 5, 2).Value = row[1]; //客戶簡稱
                    wsheet_MPT.Cell(i + 5, 3).Value = row[2]; //銷貨日期
                    wsheet_MPT.Cell(i + 5, 4).Style.NumberFormat.Format = "@";
                    wsheet_MPT.Cell(i + 5, 4).Value = row[3]; //單別
                    wsheet_MPT.Cell(i + 5, 5).Style.NumberFormat.Format = "@";
                    wsheet_MPT.Cell(i + 5, 5).Value = row[4]; //單號
                    wsheet_MPT.Cell(i + 5, 6).Value = row[5]; //批號
                    wsheet_MPT.Cell(i + 5, 7).Value = row[6]; //品號

                    if (row[6].ToString().Substring(0, 3) == "MPT")
                    {
                        sum_pMPT += "N" + (i + 5) + "+";
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDD7EE");
                    }

                    wsheet_MPT.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                    wsheet_MPT.Cell(i + 5, 8).Value = row[7]; //數量
                    wsheet_MPT.Cell(i + 5, 9).Style.NumberFormat.Format = "#,##0";
                    wsheet_MPT.Cell(i + 5, 9).Value = row[8]; //原幣未稅金額
                    wsheet_MPT.Cell(i + 5, 10).Style.NumberFormat.Format = "#,##0";
                    wsheet_MPT.Cell(i + 5, 10).Value = row[9]; //原幣稅額
                    wsheet_MPT.Cell(i + 5, 11).Style.NumberFormat.Format = "#,##0";
                    wsheet_MPT.Cell(i + 5, 11).Value = row[10]; //原幣合計金額
                    wsheet_MPT.Cell(i + 5, 12).Value = row[11]; //幣別
                    wsheet_MPT.Cell(i + 5, 13).Style.NumberFormat.Format = "#,##0.000";
                    wsheet_MPT.Cell(i + 5, 13).Value = row[12]; //匯率
                    wsheet_MPT.Cell(i + 5, 14).Style.NumberFormat.Format = "#,##0";
                    wsheet_MPT.Cell(i + 5, 14).Value = row[13]; //本幣未稅金額
                    wsheet_MPT.Cell(i + 5, 15).Style.NumberFormat.Format = "#,##0";
                    wsheet_MPT.Cell(i + 5, 15).Value = row[14]; //本幣稅額
                    wsheet_MPT.Cell(i + 5, 16).Style.NumberFormat.Format = "#,##0";
                    wsheet_MPT.Cell(i + 5, 16).Value = row[15]; //本幣合計金額

                    cust_pMPT = row[0].ToString().Trim();

                    if ((rows_count_pMPT - 1) == dt_pMPT.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        i++;
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_MPT.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_MPT.Cell(i + 5, 1).Value = cust_pMPT;
                        wsheet_MPT.Cell(i + 5, 3).Value = "小計";
                        wsheet_MPT.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_MPT.Cell(i + 5, 8).FormulaA1 = "=sum(H" + (y + 5) + ":H" + (i + 4) + ")";
                        wsheet_MPT.Range("N" + (i + 5) + ":P" + (i + 5)).Style.NumberFormat.Format = "#,##0";
                        wsheet_MPT.Cell(i + 5, 14).FormulaA1 = "=sum(N" + (y + 5) + ":N" + (i + 4) + ")";
                        wsheet_MPT.Cell(i + 5, 15).FormulaA1 = "=sum(O" + (y + 5) + ":O" + (i + 4) + ")";
                        wsheet_MPT.Cell(i + 5, 16).FormulaA1 = "=sum(P" + (y + 5) + ":P" + (i + 4) + ")";

                        y = i + 1;

                        wsheet_MPT.Cell("B1").Value = "調整查詢";
                        wsheet_MPT.Cell("B2").Value = txt_date_s2.Text.ToString() + "~" + txt_date_e2.Text.ToString();
                        wsheet_MPT.Cell("B3").Value = createday;
                        wsheet_MPT.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_MPT.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_MPT.Cell(i + 6, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_MPT.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_MPT.Cell(i + 6, 3).Value = "總計";
                        wsheet_MPT.Cell(i + 6, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_MPT.Cell(i + 6, 8).FormulaA1 = "=SUMIF(C:C,\"小計\",H:H)";
                        wsheet_MPT.Range("N" + (i + 6) + ":P" + (i + 6)).Style.NumberFormat.Format = "#,##0";
                        wsheet_MPT.Cell(i + 6, 14).FormulaA1 = "=SUMIF(C:C,\"小計\",N:N)";
                        wsheet_MPT.Cell(i + 6, 15).FormulaA1 = "=SUMIF(C:C,\"小計\",O:O)";
                        wsheet_MPT.Cell(i + 6, 16).FormulaA1 = "=SUMIF(C:C,\"小計\",P:P)";

                        wsheet_MPT.Cell(i + 10, 13).Value = "折讓";
                        wsheet_MPT.Cell(i + 10, 14).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_MPT.Cell(i + 10, 14).FormulaA1 = "=-折讓_光阻!N" + (j + 13);

                        wsheet_MPT.Range("M" + (i + 11) + ":N" + (i + 11)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_MPT.Cell(i + 11, 13).Value = "合計";
                        wsheet_MPT.Cell(i + 11, 14).Style.NumberFormat.Format = "#,##0";
                        wsheet_MPT.Cell(i + 11, 14).FormulaA1 = "=N" + (i + 6) + "+N" + (i + 10);

                        //調整
                        wsheet_MPT.Range("F" + (i + 14) + ":N" + (i + 17)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_MPT.Range("F" + (i + 14) + ":N" + (i + 17)).Style.Font.FontSize = 14;
                        wsheet_MPT.Range("K" + (i + 15) + ":N" + (i + 17)).Style.NumberFormat.Format = "#,##0";

                        wsheet_MPT.Cell(i + 14, 5).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_MPT.Cell(i + 14, 5).Style.Font.FontSize = 14;
                        wsheet_MPT.Cell(i + 14, 5).Value = "調整前";

                        wsheet_MPT.Range("F" + (i + 14) + ":N" + (i + 17)).Style.Border.OutsideBorder = XLBorderStyleValues.Double;
                        wsheet_MPT.Range("H" + (i + 14) + ":H" + (i + 16)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsheet_MPT.Cell(i + 14, 11).Value = "銷貨";
                        wsheet_MPT.Cell(i + 14, 12).Value = "6%權利金";
                        wsheet_MPT.Cell(i + 14, 13).Value = "10%稅金";
                        wsheet_MPT.Cell(i + 14, 14).Value = "淨額";

                        wsheet_MPT.Cell(i + 15, 6).Value = "Royalty";
                        wsheet_MPT.Cell(i + 15, 8).Value = "光阻";
                        wsheet_MPT.Cell(i + 15, 10).Value = "6%";
                        wsheet_MPT.Cell(i + 15, 11).FormulaA1 = "N" + (i + 6) + "+N" + (i + 10) + "-K" + (i + 16);
                        wsheet_MPT.Cell(i + 15, 12).FormulaA1 = "= ROUND(K" + (i + 15) + "* 0.06, 0)";
                        wsheet_MPT.Cell(i + 15, 13).FormulaA1 = "= ROUND(L" + (i + 15) + "* 0.1, 0)";
                        wsheet_MPT.Cell(i + 15, 14).FormulaA1 = "= L" + (i + 15) + "-M" + (i + 15);

                        wsheet_MPT.Cell(i + 16, 8).Value = "MPT";
                        wsheet_MPT.Cell(i + 16, 10).Value = "2%";

                        if (sum_pMPT.Length != 0){
                            wsheet_MPT.Cell(i + 16, 11).FormulaA1 = sum_pMPT.Substring(0, sum_pMPT.Length - 1);
                        }

                        wsheet_MPT.Cell(i + 16, 12).FormulaA1 = "= ROUND(K" + (i + 16) + "* 0.02, 0)";
                        wsheet_MPT.Cell(i + 16, 13).FormulaA1 = "= ROUND(L" + (i + 16) + "* 0.1, 0)";
                        wsheet_MPT.Cell(i + 16, 14).FormulaA1 = "= L" + (i + 16) + "-M" + (i + 16);

                        wsheet_MPT.Cell(i + 17, 8).Value = "光阻+MPT";
                        wsheet_MPT.Cell(i + 17, 11).FormulaA1 = "=sum(K" + (i + 15) + ":K" + (i + 16) + ")";
                        wsheet_MPT.Cell(i + 17, 12).FormulaA1 = "=sum(L" + (i + 15) + ":L" + (i + 16) + ")";
                        wsheet_MPT.Cell(i + 17, 13).FormulaA1 = "=sum(M" + (i + 15) + ":M" + (i + 16) + ")";
                        wsheet_MPT.Cell(i + 17, 14).FormulaA1 = "=sum(N" + (i + 15) + ":N" + (i + 16) + ")";

                                wsheet_MPT.Cell(i + 20, 13).Value = "總帳調整";
                                wsheet_MPT.Cell(i + 20, 14).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";

                        switch (str_date_e2_m)
                        {
                            case "03":
                            case "06":
                            case "09":
                            case "12":
                                wsheet_MPT.Cell(i + 20, 14).FormulaA1 = String.Format(
                                    @"=-SUM(SUMIFS(N:N,A:A,""CHOT"",C:C,"">={0}15""), SUMIFS(N:N,A:A,{{""CPCF"",""CSOT"",""WCSOT""}},C:C,"">={1}21""))"
                                    ,str_date_e2_15.Substring(0, 6),str_date_e2_21.Substring(0, 6));
                                break;
                          
                            case "01":
                            case "04":
                            case "07":
                            case "10":
                                wsheet_MPT.Cell(i + 20, 14).FormulaA1 = String.Format(
                                    @"=SUM(SUMIFS(N:N,A:A,""CHOT"",C:C,""<{0}01""), SUMIFS(N:N,A:A,{{""CPCF"",""CSOT"",""WCSOT""}},C:C,""<{1}01""))"
                                    , str_date_e2_15.Substring(0, 6), str_date_e2_21.Substring(0, 6));
                                break;

                            default:
                                
                                break;
                        }
                                wsheet_MPT.Cell(i + 21, 13).Value = "折讓";
                                wsheet_MPT.Cell(i + 21, 14).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                                wsheet_MPT.Cell(i + 21, 14).FormulaA1 = "=-折讓_光阻!N" + (j + 13);

                                wsheet_MPT.Range("M" + (i + 22) + ":N" + (i + 22)).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFBD45");
                                wsheet_MPT.Cell(i + 22, 13).Value = "合計";
                                wsheet_MPT.Cell(i + 22, 14).Style.NumberFormat.Format = "#,##0";
                                wsheet_MPT.Cell(i + 22, 14).FormulaA1 = "=N" + (i + 6) + "+N" + (i + 20) + "+N" + (i + 21);

                        wsheet_MPT.Range("F" + (i + 24) + ":N" + (i + 27)).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFBD45");
                        wsheet_MPT.Range("F" + (i + 24) + ":N" + (i + 27)).Style.Font.FontSize = 14;
                        wsheet_MPT.Range("K" + (i + 25) + ":N" + (i + 27)).Style.NumberFormat.Format = "#,##0";

                        wsheet_MPT.Cell(i + 24, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFBD45");
                        wsheet_MPT.Cell(i + 24, 5).Style.Font.FontSize = 14;
                        wsheet_MPT.Cell(i + 24, 5).Value = "調整後";

                        wsheet_MPT.Range("F" + (i + 24) + ":N" + (i + 27)).Style.Border.OutsideBorder = XLBorderStyleValues.Double;
                        wsheet_MPT.Range("H" + (i + 24) + ":H" + (i + 26)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsheet_MPT.Cell(i + 24, 11).Value = "銷貨";
                        wsheet_MPT.Cell(i + 24, 12).Value = "6%權利金";
                        wsheet_MPT.Cell(i + 24, 13).Value = "10%稅金";
                        wsheet_MPT.Cell(i + 24, 14).Value = "淨額";

                        wsheet_MPT.Cell(i + 25, 6).Value = "Royalty";
                        wsheet_MPT.Cell(i + 25, 8).Value = "光阻";
                        wsheet_MPT.Cell(i + 25, 10).Value = "6%";
                        wsheet_MPT.Cell(i + 25, 11).FormulaA1 = "N" + (i + 6) + "+N" + (i + 20) + "+N" + (i + 21) + "-K" + (i + 26);
                        wsheet_MPT.Cell(i + 25, 12).FormulaA1 = "= ROUND(K" + (i + 25) + "* 0.06, 0)";
                        wsheet_MPT.Cell(i + 25, 13).FormulaA1 = "= ROUND(L" + (i + 25) + "* 0.1, 0)";
                        wsheet_MPT.Cell(i + 25, 14).FormulaA1 = "= L" + (i + 25) + "-M" + (i + 25);

                        wsheet_MPT.Cell(i + 26, 8).Value = "MPT";
                        wsheet_MPT.Cell(i + 26, 10).Value = "2%";

                        if (sum_pMPT.Length != 0){
                            wsheet_MPT.Cell(i + 26, 11).FormulaA1 = sum_pMPT.Substring(0, sum_pMPT.Length - 1);
                        }

                        wsheet_MPT.Cell(i + 26, 12).FormulaA1 = "= ROUND(K" + (i + 26) + "* 0.02, 0)";
                        wsheet_MPT.Cell(i + 26, 13).FormulaA1 = "= ROUND(L" + (i + 26) + "* 0.1, 0)";
                        wsheet_MPT.Cell(i + 26, 14).FormulaA1 = "= L" + (i + 26) + "-M" + (i + 26);

                        wsheet_MPT.Cell(i + 27, 8).Value = "光阻+MPT";
                        wsheet_MPT.Cell(i + 27, 11).FormulaA1 = "=sum(K" + (i + 25) + ":K" + (i + 26) + ")";
                        wsheet_MPT.Cell(i + 27, 12).FormulaA1 = "=sum(L" + (i + 25) + ":L" + (i + 26) + ")";
                        wsheet_MPT.Cell(i + 27, 13).FormulaA1 = "=sum(M" + (i + 25) + ":M" + (i + 26) + ")";
                        wsheet_MPT.Cell(i + 27, 14).FormulaA1 = "=sum(N" + (i + 25) + ":N" + (i + 26) + ")";
                    }
                    i++;
                }
                //worksheet.Columns().AdjustToContents();
                //worksheet2.Columns().AdjustToContents();

                wsheet_MPT.Position = 1;

                wsheet_MPT.Column("A").Width = 7;
                wsheet_MPT.Column("B").Width = 10;
                wsheet_MPT.Column("C").Width = 8;
                wsheet_MPT.Column("D").Width = 4;
                wsheet_MPT.Column("E").Width = 6.5;
                wsheet_MPT.Column("F").Width = 7.5;
                wsheet_MPT.Column("G").Width = 16;
                wsheet_MPT.Column("H").Width = 11;
                wsheet_MPT.Column("I").Width = 10;
                wsheet_MPT.Column("J").Width = 7;


                wsheet_MPT.Column("K").Width = 14;
                wsheet_MPT.Column("L").Width = 14;
                wsheet_MPT.Column("M").Width = 14;
                wsheet_MPT.Column("N").Width = 14;

                save_as_MPT = txt_path.Text.ToString().Trim() + @"\\銷貨明細(權利金)光阻_調整" + txt_date_e2.Text.ToString().Substring(0, 6) + "_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";

                wb_MPT.SaveAs(save_as_MPT);

                //打开文件
                //System.Diagnostics.Process.Start(save_as_MPT);
            }

            //保存成文件 日本凸版
            using (XLWorkbook wb_JP = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel))
                {
                    var ws = templateWB.Worksheet(1);
                    var ws2 = templateWB.Worksheet(3);
                    var ws3 = templateWB.Worksheet(4);

                    ws.CopyTo(wb_JP, "權利金_日本凸版");
                    ws2.CopyTo(wb_JP, "折讓_日本凸版");
                    ws3.CopyTo(wb_JP, "手續費");
                }

                //var worksheet2 = wb_JP.Worksheets.Add("折讓_光阻");
                var wsheet_JP = wb_JP.Worksheet("權利金_日本凸版");
                var wsheet_dcJP = wb_JP.Worksheet("折讓_日本凸版");
                var wsheet_hJP = wb_JP.Worksheet("手續費");

                //== 手續費-日本凸版 ================================================================
                int rows_count_hJP = dt_hJP.Rows.Count;
                int k = 0; int z = 0; int q = 0; int r = 0;
                string cust_hJP = "";
                string custid_hJP = "";
                int[] custidnum_hJP = new int[20];
                string[] custidname_hJP = new string[20];
                int[] num_hJP = new int[20];

                if (rows_count_hJP == 0)
                {
                    wsheet_hJP.Cell("B2").Value = txt_date_s2.Text.ToString() + "~" + txt_date_e2.Text.ToString();
                    wsheet_hJP.Cell("B3").Value = createday;

                    wsheet_hJP.Cell(k + 9, 7).Value = "手續費";

                    wsheet_hJP.Range("G" + (k + 10) + ":H" + (k + 10)).Style.Fill.BackgroundColor = XLColor.Yellow;
                    wsheet_hJP.Cell(k + 10, 7).Value = "合計";
                    wsheet_hJP.Cell(k + 10, 8).Value = "0";

                    wsheet_hJP.Range("F" + (k + 13) + ":H" + (k + 15)).Style.Fill.BackgroundColor = XLColor.Yellow;
                    wsheet_hJP.Range("H" + (k + 13) + ":H" + (k + 15)).Style.NumberFormat.Format = "#,##0";
                    wsheet_hJP.Range("F" + (k + 13) + ":F" + (k + 14)).Style.Font.FontSize = 10;


                    wsheet_hJP.Range("H" + (k + 13) + ":H" + (k + 15)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    wsheet_hJP.Range("F" + (k + 14) + ":H" + (k + 14)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    wsheet_hJP.Range("F" + (k + 15) + ":H" + (k + 15)).Style.Border.BottomBorder = XLBorderStyleValues.Double;

                    wsheet_hJP.Cell(k + 13, 6).Value = "手續費(2%)-CSOT、WCSOT、CPCF、AU-L6K、AU-TC、AU-TY";
                    wsheet_hJP.Cell(k + 13, 8).FormulaA1 = "=SUM(SUMIFS(H:H,A:A,{\"CSOT\",\"WCSOT\",\"CPCF\",\"AU-L6K\",\"AU-TC\",\"AU-TY\"},B:B,\"小計\"))";
                    num_hJP[0] = k + 13;

                    wsheet_hJP.Cell(k + 14, 6).Value = "手續費(1%)-BOE、CPD、HKC、HKC-H2、HKC-H4、HKC-H5、CHOT、HSD、CCPD";
                    wsheet_hJP.Cell(k + 14, 8).FormulaA1 = "=SUM(SUMIFS(H:H,A:A,{\"BOE\",\"CPD\",\"HKC\",\"HKC-H2\",\"HKC-H4\",\"HKC-H5\",\"CHOT\",\"HSD\",\"CCPD\"},B:B,\"小計\"))";
                    num_hJP[1] = k + 14;

                    wsheet_hJP.Cell(k + 15, 6).Value = "合計";
                    wsheet_hJP.Cell(k + 15, 8).FormulaA1 = "=sum(H" + (k + 13) + ":H" + (k + 14) + ")";
                }
                else
                {
                    foreach (DataRow row in dt_hJP.Rows)
                    {
                        if (cust_hJP.ToString() != "" && row[0].ToString() != cust_hJP.ToString())
                        {
                            wsheet_hJP.Range("A" + (k + 5) + ":J" + (k + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Range("A" + (k + 5) + ":J" + (k + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Range("A" + (k + 5) + ":J" + (k + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                            wsheet_hJP.Cell(k + 5, 1).Value = custid_hJP;
                            wsheet_hJP.Cell(k + 5, 2).Value = "小計";
                            wsheet_hJP.Cell(k + 5, 8).FormulaA1 = "=sum(H" + (z + 5) + ":H" + (k + 4) + ")";

                            custidname_hJP[q] = custid_hJP;

                            q += 1;
                            z = k + 1;

                            k++;
                        }

                        wsheet_hJP.Cell(k + 5, 1).Value = row[0]; //客戶名稱
                        wsheet_hJP.Cell(k + 5, 2).Value = row[1]; //科目編號
                        wsheet_hJP.Cell(k + 5, 3).Value = row[2]; //科目名稱
                        wsheet_hJP.Cell(k + 5, 4).Value = row[3]; //傳票日期
                        wsheet_hJP.Cell(k + 5, 5).Value = row[4]; //傳票編號
                        wsheet_hJP.Cell(k + 5, 6).Value = row[5]; //摘要
                        wsheet_hJP.Cell(k + 5, 7).Value = row[6]; //手續費率
                        wsheet_hJP.Cell(k + 5, 8).Style.NumberFormat.Format = "#,##0";
                        wsheet_hJP.Cell(k + 5, 8).Value = row[7]; //借方金額
                        wsheet_hJP.Cell(k + 5, 9).Style.NumberFormat.Format = "#,##0";
                        wsheet_hJP.Cell(k + 5, 9).Value = row[8]; //貸方金額
                        wsheet_hJP.Cell(k + 5, 10).Value = row[9]; //借貸

                        cust_hJP = row[0].ToString().Trim();
                        custid_hJP = row[0].ToString().Trim();

                        if ((rows_count_hJP - 1) == dt_hJP.Rows.IndexOf(row)) //資料列結尾運算
                        {
                            k++;
                            wsheet_hJP.Range("A" + (k + 5) + ":J" + (k + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Range("A" + (k + 5) + ":J" + (k + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Range("A" + (k + 5) + ":J" + (k + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                            wsheet_hJP.Cell(k + 5, 1).Value = custid_hJP;
                            wsheet_hJP.Cell(k + 5, 2).Value = "小計";
                            wsheet_hJP.Cell(k + 5, 8).FormulaA1 = "=sum(H" + (z + 5) + ":H" + (k + 4) + ")";

                            custidname_hJP[q] = custid_hJP;
                            //custidnum_dcJP[n] = (k + 5);

                            //sum_dcJP += "N" + (k + 5);
                            //x = j + 1;

                            z++;

                            wsheet_hJP.Cell("B2").Value = txt_date_s2.Text.ToString() + "~" + txt_date_e2.Text.ToString();
                            wsheet_hJP.Cell("B3").Value = createday;
                            wsheet_hJP.Range("A" + (k + 6) + ":J" + (k + 6)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Range("A" + (k + 6) + ":J" + (k + 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Cell(k + 6, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                            wsheet_hJP.Range("A" + (k + 6) + ":J" + (k + 6)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                            wsheet_hJP.Cell(k + 6, 2).Value = "總計";
                            wsheet_hJP.Cell(k + 6, 8).Style.NumberFormat.Format = "#,##0.00";
                            wsheet_hJP.Cell(k + 6, 8).FormulaA1 = "=SUMIF(B:B,\"小計\",H:H)";

                            wsheet_hJP.Cell(k + 9, 7).Value = "手續費";

                            for (int num = 0; num <= q; num++)
                            {
                                wsheet_hJP.Cell(k + 10, 7).Value = custidname_hJP[num];
                                wsheet_hJP.Cell(k + 10, 8).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                                wsheet_hJP.Cell(k + 10, 8).FormulaA1 = "=SUMIFS(H:H,A:A,\"" + custidname_hJP[num] + "\",B:B,\"小計\")";
                                k++;
                                r++;
                            }

                            wsheet_hJP.Cell(k + 10, 7).Value = "合計";
                            wsheet_hJP.Cell(k + 10, 8).Style.NumberFormat.Format = "#,##0";

                            wsheet_hJP.Range("G" + (k + 10) + ":H" + (k + 10)).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wsheet_hJP.Cell(k + 10, 8).FormulaA1 = "=SUM(H" + (k + 10 - r) + ":H" + (k + 9) + ")";

                            wsheet_hJP.Range("F" + (k + 13) + ":H" + (k + 15)).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wsheet_hJP.Range("H" + (k + 13) + ":H" + (k + 15)).Style.NumberFormat.Format = "#,##0";
                            wsheet_hJP.Range("F" + (k + 13) + ":F" + (k + 14)).Style.Font.FontSize = 10;


                            wsheet_hJP.Range("H" + (k + 13) + ":H" + (k + 15)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            wsheet_hJP.Range("F" + (k + 14) + ":H" + (k + 14)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            wsheet_hJP.Range("F" + (k + 15) + ":H" + (k + 15)).Style.Border.BottomBorder = XLBorderStyleValues.Double;

                            wsheet_hJP.Cell(k + 13, 6).Value = "手續費(2%)-CSOT、WCSOT、CPCF、AU-L6K、AU-TC、AU-TY";
                            wsheet_hJP.Cell(k + 13, 8).FormulaA1 = "=SUM(SUMIFS(H:H,A:A,{\"CSOT\",\"WCSOT\",\"CPCF\",\"AU-L6K\",\"AU-TC\",\"AU-TY\"},B:B,\"小計\"))";
                            num_hJP[0] = k + 13;

                            wsheet_hJP.Cell(k + 14, 6).Value = "手續費(1%)-BOE、CPD、HKC、HKC-H2、HKC-H4、HKC-H5、CHOT、HSD、CCPD";
                            wsheet_hJP.Cell(k + 14, 8).FormulaA1 = "=SUM(SUMIFS(H:H,A:A,{\"BOE\",\"CPD\",\"HKC\",\"HKC-H2\",\"HKC-H4\",\"HKC-H5\",\"CHOT\",\"HSD\",\"CCPD\"},B:B,\"小計\"))";
                            num_hJP[1] = k + 14;

                            wsheet_hJP.Cell(k + 15, 6).Value = "合計";
                            wsheet_hJP.Cell(k + 15, 8).FormulaA1 = "=sum(H" + (k + 13) + ":H" + (k + 14) + ")";
                        }
                        k++;
                    }
                }

                //== 折讓-日本凸版 ================================================================
                int rows_count_dcJP = dt_dcJP.Rows.Count;
                int j = 0; int x = 0; int n = 0;
                string cust_dcJP = "";
                string custid_dcJP = "";
                int[] custidnum_dcJP = new int[20];
                string[] custidname_dcJP = new string[20];
                int[] num_dcJP = new int[20];

                foreach (DataRow row in dt_dcJP.Rows)
                {
                    if (cust_dcJP.ToString() != "" && row[0].ToString() != cust_dcJP.ToString())
                    {
                        wsheet_dcJP.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_dcJP.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_dcJP.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_dcJP.Cell(j + 7, 1).Value = custid_dcJP;
                        wsheet_dcJP.Cell(j + 7, 2).Value = "小計";
                        wsheet_dcJP.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 7, 11).FormulaA1 = "=sum(K" + (x + 7) + ":K" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 12).FormulaA1 = "=sum(L" + (x + 7) + ":L" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 13).FormulaA1 = "=sum(M" + (x + 7) + ":M" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 14).FormulaA1 = "=sum(N" + (x + 7) + ":N" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 15).FormulaA1 = "=sum(O" + (x + 7) + ":O" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 7, 17).FormulaA1 = "=sum(Q" + (x + 7) + ":Q" + (j + 6) + ")";

                        switch (custid_dcJP)
                        {
                            case "CSOT":
                            case "WCSOT":
                            case "CPCF":
                            case "AU-L6K":
                            case "AU-TC":
                            case "AU-TY":
                            case "BOE":
                            case "CPD":
                            case "HKC":
                            case "HKC-H2":
                            case "HKC-H4":
                            case "HKC-H5":
                            case "CCPD":
                            case "CHOT":
                            case "HSD":
                                custidname_dcJP[n] = custid_dcJP;
                                custidnum_dcJP[n] = (j + 7);
                                break;
                            default:
                                break;
                        }

                        //sum_dcJP += "N" + (j + 7) + "+";

                        n += 1;
                        x = j + 1;
                        j++;
                    }

                    if (row[0].ToString() != cust_dcJP.ToString())
                    {
                        wsheet_dcJP.Cell(j + 7, 1).Value = row[0]; //客戶
                    }

                    wsheet_dcJP.Cell(j + 7, 2).Value = row[1]; //銷貨日期
                    wsheet_dcJP.Cell(j + 7, 3).Style.NumberFormat.Format = "@";
                    wsheet_dcJP.Cell(j + 7, 3).Value = row[2]; //折讓單別
                    wsheet_dcJP.Cell(j + 7, 4).Style.NumberFormat.Format = "@";
                    wsheet_dcJP.Cell(j + 7, 4).Value = row[3]; //折讓單號
                    wsheet_dcJP.Cell(j + 7, 5).Style.NumberFormat.Format = "@";
                    wsheet_dcJP.Cell(j + 7, 5).Value = row[4]; //客戶單號
                    wsheet_dcJP.Cell(j + 7, 6).Value = row[5]; //品號
                    wsheet_dcJP.Cell(j + 7, 7).Value = row[6]; //批號
                    wsheet_dcJP.Cell(j + 7, 8).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 8).Value = row[7]; //銷貨單價
                    wsheet_dcJP.Cell(j + 7, 9).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 9).Value = row[8]; //新單價
                    wsheet_dcJP.Cell(j + 7, 10).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 10).Value = row[9]; //折讓差

                    wsheet_dcJP.Cell(j + 7, 11).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 11).Value = row[10]; //銷貨數量
                    wsheet_dcJP.Cell(j + 7, 12).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 12).Value = row[11]; //折讓金額
                    wsheet_dcJP.Cell(j + 7, 13).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 13).Value = row[12]; //折讓稅額
                    wsheet_dcJP.Cell(j + 7, 14).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 14).Value = row[13]; //台幣金額
                    wsheet_dcJP.Cell(j + 7, 15).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 15).Value = row[14]; //台幣稅額
                    wsheet_dcJP.Cell(j + 7, 16).Style.NumberFormat.Format = "#,##0.0000";
                    wsheet_dcJP.Cell(j + 7, 16).Value = row[15]; //匯率
                    wsheet_dcJP.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                    wsheet_dcJP.Cell(j + 7, 17).Value = row[16]; //台幣合計
                    wsheet_dcJP.Cell(j + 7, 18).Value = row[17]; //發票號碼
                    wsheet_dcJP.Cell(j + 7, 19).Value = row[18]; //廠別

                    wsheet_dcJP.Cell(j + 7, 20).Value = row[19]; //銷退日

                    cust_dcJP = row[0].ToString().Trim();
                    custid_dcJP = row[0].ToString().Trim();

                    if ((rows_count_dcJP - 1) == dt_dcJP.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        j++;
                        wsheet_dcJP.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_dcJP.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_dcJP.Range("A" + (j + 7) + ":S" + (j + 7)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_dcJP.Cell(j + 7, 1).Value = custid_dcJP;
                        wsheet_dcJP.Cell(j + 7, 2).Value = "小計";
                        wsheet_dcJP.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 7, 11).FormulaA1 = "=sum(K" + (x + 7) + ":K" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 12).FormulaA1 = "=sum(L" + (x + 7) + ":L" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 13).FormulaA1 = "=sum(M" + (x + 7) + ":M" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 14).FormulaA1 = "=sum(N" + (x + 7) + ":N" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 15).FormulaA1 = "=sum(O" + (x + 7) + ":O" + (j + 6) + ")";
                        wsheet_dcJP.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 7, 17).FormulaA1 = "=sum(Q" + (x + 7) + ":Q" + (j + 6) + ")";

                        custidname_dcJP[n] = custid_dcJP;
                        custidnum_dcJP[n] = (j + 7);

                        //sum_dcJP += "N" + (j + 7);
                        x = j + 1;

                        j++;

                        wsheet_dcJP.Cell("B4").Value = txt_date_s2.Text.ToString() + " ~ " + txt_date_e2.Text.ToString();
                        wsheet_dcJP.Cell("B5").Value = createday;

                        wsheet_dcJP.Cell(j + 7, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_dcJP.Range("A" + (j + 7) + ":T" + (j + 7)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_dcJP.Cell(j + 7, 3).Value = "總計";
                        wsheet_dcJP.Range("K" + (j + 7) + ":O" + (j + 7)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 7, 11).FormulaA1 = "=SUMIF(B:B,\"小計\",K:K)";
                        wsheet_dcJP.Cell(j + 7, 12).FormulaA1 = "=SUMIF(B:B,\"小計\",L:L)";
                        wsheet_dcJP.Cell(j + 7, 13).FormulaA1 = "=SUMIF(B:B,\"小計\",M:M)";
                        wsheet_dcJP.Cell(j + 7, 14).FormulaA1 = "=SUMIF(B:B,\"小計\",N:N)";
                        wsheet_dcJP.Cell(j + 7, 15).FormulaA1 = "=SUMIF(B:B,\"小計\",O:O)";
                        wsheet_dcJP.Cell(j + 7, 17).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 7, 17).FormulaA1 = "=SUMIF(B:B,\"小計\",Q:Q)";

                        j = j + 2;
                        wsheet_dcJP.Range("N" + (j + 10) + ":N" + (j + 15)).Style.NumberFormat.Format = "#,##0";
                        wsheet_dcJP.Cell(j + 9, 13).Value = "UV製品-折讓";
                        wsheet_dcJP.Cell(j + 10, 13).Value = " 製品-折讓2%";
                        wsheet_dcJP.Cell(j + 10, 14).FormulaA1 = "=SUM(SUMIFS(N:N,A:A,{\"CSOT\",\"WCSOT\",\"CPCF\",\"AU-L6K\",\"AU-TC\",\"AU-TY\"},B:B,\"小計\"))";
                        num_dcJP[0] = (j + 10);

                        wsheet_dcJP.Cell(j + 11, 13).Value = " 製品-折讓1%";
                        wsheet_dcJP.Cell(j + 11, 14).FormulaA1 = "=SUM(SUMIFS(N:N,A:A,{\"BOE\",\"CPD\",\"HKC\",\"HKC-H2\",\"HKC-H4\",\"HKC-H5\",\"CHOT\",\"HSD\",\"CCPD\"},B:B,\"小計\"))";
                        num_dcJP[1] = (j + 11);

                        wsheet_dcJP.Cell(j + 12, 13).Value = " 製品-退回";
                        wsheet_dcJP.Cell(j + 13, 13).Value = " 商品-折讓";
                        wsheet_dcJP.Cell(j + 14, 13).Value = " 商品-退回";
                        wsheet_dcJP.Cell(j + 15, 13).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_dcJP.Range("M" + (j + 15) + ":N" + (j + 15)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_dcJP.Cell(j + 15, 13).Value = "合計";
                        wsheet_dcJP.Cell(j + 15, 14).FormulaA1 = "=sum(N" + (j + 10) + ":N" + (j + 14) + ")";
                    }
                    j++;
                }

                //== 權利金-日本凸版 ================================================================
                int rows_count_pJP = dt_pJP.Rows.Count;
                int i = 0; int y = 0; int m = 0; int p = 0;
                string cust_pJP = "";
                string custid_pJP = "";
                string[] custidname_pJP = new string[20];
                int[] sumrow_pJP = new int[20];

                foreach (DataRow row in dt_pJP.Rows)
                {
                    if (cust_pJP.ToString() != "" && row[0].ToString() != cust_pJP.ToString())
                    {
                        wsheet_JP.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_JP.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_JP.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_JP.Cell(i + 5, 1).Value = custid_pJP;
                        wsheet_JP.Cell(i + 5, 3).Value = "小計";
                        wsheet_JP.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_JP.Cell(i + 5, 8).FormulaA1 = "=sum(H" + (y + 5) + ":H" + (i + 4) + ")";
                        wsheet_JP.Range("N" + (i + 5) + ":P" + (i + 5)).Style.NumberFormat.Format = "#,##0";
                        wsheet_JP.Cell(i + 5, 14).FormulaA1 = "=sum(N" + (y + 5) + ":N" + (i + 4) + ")";
                        wsheet_JP.Cell(i + 5, 15).FormulaA1 = "=sum(O" + (y + 5) + ":O" + (i + 4) + ")";
                        wsheet_JP.Cell(i + 5, 16).FormulaA1 = "=sum(P" + (y + 5) + ":P" + (i + 4) + ")";

                        custidname_pJP[m] = custid_pJP.ToString();
                        m += 1;
                        y = i + 1;
                        i++;
                    }

                    //填入excel欄位值
                    wsheet_JP.Cell(i + 5, 1).Value = row[0]; //客戶代號
                    wsheet_JP.Cell(i + 5, 2).Value = row[1]; //客戶簡稱
                    wsheet_JP.Cell(i + 5, 3).Value = row[2]; //銷貨日期
                    wsheet_JP.Cell(i + 5, 4).Style.NumberFormat.Format = "@";
                    wsheet_JP.Cell(i + 5, 4).Value = row[3]; //單別
                    wsheet_JP.Cell(i + 5, 5).Style.NumberFormat.Format = "@";
                    wsheet_JP.Cell(i + 5, 5).Value = row[4]; //單號
                    wsheet_JP.Cell(i + 5, 6).Value = row[5]; //批號
                    wsheet_JP.Cell(i + 5, 7).Value = row[6]; //品號
                    wsheet_JP.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                    wsheet_JP.Cell(i + 5, 8).Value = row[7]; //數量
                    wsheet_JP.Cell(i + 5, 9).Style.NumberFormat.Format = "#,##0";
                    wsheet_JP.Cell(i + 5, 9).Value = row[8]; //原幣未稅金額
                    wsheet_JP.Cell(i + 5, 10).Style.NumberFormat.Format = "#,##0";
                    wsheet_JP.Cell(i + 5, 10).Value = row[9]; //原幣稅額
                    wsheet_JP.Cell(i + 5, 11).Style.NumberFormat.Format = "#,##0";
                    wsheet_JP.Cell(i + 5, 11).Value = row[10]; //原幣合計金額
                    wsheet_JP.Cell(i + 5, 12).Value = row[11]; //幣別
                    wsheet_JP.Cell(i + 5, 13).Style.NumberFormat.Format = "#,##0.000";
                    wsheet_JP.Cell(i + 5, 13).Value = row[12]; //匯率
                    wsheet_JP.Cell(i + 5, 14).Style.NumberFormat.Format = "#,##0";
                    wsheet_JP.Cell(i + 5, 14).Value = row[13]; //本幣未稅金額
                    wsheet_JP.Cell(i + 5, 15).Style.NumberFormat.Format = "#,##0";
                    wsheet_JP.Cell(i + 5, 15).Value = row[14]; //本幣稅額
                    wsheet_JP.Cell(i + 5, 16).Style.NumberFormat.Format = "#,##0";
                    wsheet_JP.Cell(i + 5, 16).Value = row[15]; //本幣合計金額

                    cust_pJP = row[0].ToString().Trim();
                    custid_pJP = row[0].ToString().Trim();


                    if ((rows_count_pJP - 1) == dt_pJP.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        i++;
                        wsheet_JP.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_JP.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_JP.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_JP.Cell(i + 5, 1).Value = custid_pJP;
                        wsheet_JP.Cell(i + 5, 3).Value = "小計";
                        wsheet_JP.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_JP.Cell(i + 5, 8).FormulaA1 = "=sum(H" + (y + 5) + ":H" + (i + 4) + ")";
                        wsheet_JP.Range("N" + (i + 5) + ":P" + (i + 5)).Style.NumberFormat.Format = "#,##0";
                        wsheet_JP.Cell(i + 5, 14).FormulaA1 = "=sum(N" + (y + 5) + ":N" + (i + 4) + ")";
                        wsheet_JP.Cell(i + 5, 15).FormulaA1 = "=sum(O" + (y + 5) + ":O" + (i + 4) + ")";
                        wsheet_JP.Cell(i + 5, 16).FormulaA1 = "=sum(P" + (y + 5) + ":P" + (i + 4) + ")";

                        custidname_pJP[m] = custid_pJP.ToString();
                        y = i + 1;

                        wsheet_JP.Cell("B1").Value = "調整查詢";
                        wsheet_JP.Cell("B2").Value = txt_date_s2.Text.ToString() + "~" + txt_date_e2.Text.ToString();
                        wsheet_JP.Cell("B3").Value = createday;
                        wsheet_JP.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_JP.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_JP.Cell(i + 6, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_JP.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_JP.Cell(i + 6, 3).Value = "總計";
                        wsheet_JP.Cell(i + 6, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_JP.Cell(i + 6, 8).FormulaA1 = "=SUMIF(C:C,\"小計\",H:H)";
                        sumrow_pJP[0] = i + 6;
                        wsheet_JP.Range("N" + (i + 6) + ":P" + (i + 6)).Style.NumberFormat.Format = "#,##0";
                        wsheet_JP.Cell(i + 6, 14).FormulaA1 = "=SUMIF(C:C,\"小計\",N:N)";
                        wsheet_JP.Cell(i + 6, 15).FormulaA1 = "=SUMIF(C:C,\"小計\",O:O)";
                        wsheet_JP.Cell(i + 6, 16).FormulaA1 = "=SUMIF(C:C,\"小計\",P:P)";

                        //調整 
                        wsheet_JP.Range("E" + (i + 10) + ":N" + (i + 28)).Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                        wsheet_JP.Cell(i + 10, 5).Value = "調整前";
                        wsheet_JP.Range("N" + (i + 10) + ":N" + (i + 22)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_JP.Cell(i + 10, 13).Value = "AU(AU-C5E)";
                        wsheet_JP.Cell(i + 10, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"AU\",C:C,\"小計\")";
                        sumrow_pJP[1] = i + 10;
                        wsheet_JP.Cell(i + 11, 13).Value = "AU-T(AU-C4A)";
                        wsheet_JP.Cell(i + 11, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"AU-T\",C:C,\"小計\")";

                        wsheet_JP.Cell(i + 12, 13).Value = "AU-TK(路竹)";
                        wsheet_JP.Cell(i + 12, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"AU-TK\",C:C,\"小計\")";

                        wsheet_JP.Cell(i + 13, 13).Value = "AU-TN(台南)";
                        wsheet_JP.Cell(i + 13, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"AU-TN\",C:C,\"小計\")";

                        wsheet_JP.Cell(i + 14, 13).Value = "TCE";
                        wsheet_JP.Cell(i + 14, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"TCE\",C:C,\"小計\")";

                        wsheet_JP.Cell(i + 15, 13).Value = "TCET-AU";
                        wsheet_JP.Cell(i + 15, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"TCET-AU\",C:C,\"小計\")";

                        for (int num = 0; num <= m; num++)
                        {
                            switch (custidname_pJP[num])
                            {
                                case "AU":
                                case "AU-T":
                                case "AU-TK":
                                case "AU-TN":
                                case "TCE":
                                case "TCET-AU":

                                case "CSOT":
                                case "WCSOT":
                                case "CPCF":
                                case "AU-L6K":
                                case "AU-TC":
                                case "AU-TY":
                                case "BOE":
                                case "CPD":
                                case "HKC":
                                case "HKC-H2":
                                case "HKC-H4":
                                case "HKC-H5":
                                case "CCPD":
                                case "CHOT":
                                case "HSD":
                                    break;
                                default:
                                    wsheet_JP.Cell(i + 16, 13).Value = custidname_pJP[num];
                                    wsheet_JP.Cell(i + 16, 14).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                                    wsheet_JP.Cell(i + 16, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"" + custidname_pJP[num] + "\",C:C,\"小計\")";
                                    i++;
                                    break;
                            }
                            p++;
                        }

                        wsheet_JP.Range("M" + (i + 16) + ":N" + (i + 19)).Style.Fill.BackgroundColor = XLColor.FromHtml("#F1DBDA");
                        wsheet_JP.Cell(i + 16, 13).Value = "折讓-2%";
                        wsheet_JP.Cell(i + 16, 14).FormulaA1 = "=-折讓_日本凸版!N" + num_dcJP[0];
                        sumrow_pJP[3] = i + 16;

                        wsheet_JP.Cell(i + 17, 13).Value = "折讓-1%";
                        wsheet_JP.Cell(i + 17, 14).FormulaA1 = "=-折讓_日本凸版!N" + num_dcJP[1];


                        wsheet_JP.Cell(i + 18, 13).Value = "匯款手續費-2%";
                        wsheet_JP.Cell(i + 18, 14).FormulaA1 = "=-手續費!H" + num_hJP[0];

                        wsheet_JP.Cell(i + 19, 13).Value = "匯款手續費-1%";
                        wsheet_JP.Cell(i + 19, 14).FormulaA1 = "=-手續費!H" + num_hJP[1];
                        sumrow_pJP[2] = i + 19;
                        sumrow_pJP[4] = i + 19;

                        wsheet_JP.Cell(i + 20, 13).Value = "合計";
                        wsheet_JP.Cell(i + 20, 14).Style.NumberFormat.Format = "#,##0";
                        wsheet_JP.Range("M" + (i + 20) + ":N" + (i + 20)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_JP.Cell(i + 20, 14).FormulaA1 = "=N" + sumrow_pJP[0] + "+SUM(N" + sumrow_pJP[1] + ":N" + sumrow_pJP[2] + ")";

                        wsheet_JP.Range("F" + (i + 24) + ":N" + (i + 27)).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wsheet_JP.Range("F" + (i + 24) + ":N" + (i + 27)).Style.Font.FontSize = 14;
                        wsheet_JP.Range("K" + (i + 25) + ":N" + (i + 27)).Style.NumberFormat.Format = "#,##0";
                        wsheet_JP.Range("G" + (i + 25) + ":G" + (i + 26)).Style.Font.FontSize = 10;

                        wsheet_JP.Range("F" + (i + 24) + ":N" + (i + 27)).Style.Border.OutsideBorder = XLBorderStyleValues.Double;

                        wsheet_JP.Range("H" + (i + 24) + ":H" + (i + 26)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsheet_JP.Cell(i + 24, 11).Value = "銷貨";
                        wsheet_JP.Cell(i + 24, 12).Value = "權利金";
                        wsheet_JP.Cell(i + 24, 13).Value = "10%稅金";
                        wsheet_JP.Cell(i + 24, 14).Value = "淨額";

                        wsheet_JP.Cell(i + 25, 6).Value = "Royalty";
                        wsheet_JP.Cell(i + 25, 7).Value = "CSOT、WCSOT、CPCF、AU-L6K、AU-TC、AU-TY";
                        wsheet_JP.Cell(i + 25, 10).Value = "2%";
                        wsheet_JP.Cell(i + 25, 11).FormulaA1 = "=SUM(SUMIFS(N:N,A:A,{\"CSOT\",\"WCSOT\",\"CPCF\",\"AU-L6K\",\"AU-TC\",\"AU-TY\"},C:C,\"小計\"))" + str_enter +
                                                               "+SUM(SUMIFS(N" + sumrow_pJP[3] + ":N" + sumrow_pJP[4] + ",M" + sumrow_pJP[3] + ":M" + sumrow_pJP[4] + str_enter +
                                                               ",{\"折讓-2%\",\"匯款手續費-2%\"}))";

                        wsheet_JP.Cell(i + 25, 12).FormulaA1 = "= ROUND(K" + (i + 25) + "* 0.02, 0)";
                        wsheet_JP.Cell(i + 25, 13).FormulaA1 = "= ROUND(L" + (i + 25) + "* 0.1, 0)";
                        wsheet_JP.Cell(i + 25, 14).FormulaA1 = "= L" + (i + 25) + "-M" + (i + 25);

                        wsheet_JP.Cell(i + 26, 7).Value = "BOE、CPD、HKC、HKC-H2、HKC-H4、HKC-H5、CHOT、HSD、CCPD";
                        wsheet_JP.Cell(i + 26, 10).Value = "1%";
                        wsheet_JP.Cell(i + 26, 11).FormulaA1 = "=SUM(SUMIFS(N:N,A:A,{\"BOE\",\"CPD\",\"HKC\",\"HKC-H2\",\"HKC-H4\",\"HKC-H5\",\"CHOT\",\"HSD\",\"CCPD\"},C:C,\"小計\"))" + str_enter +
                                                               "+SUM(SUMIFS(N" + sumrow_pJP[3] + ":N" + sumrow_pJP[4] + ",M" + sumrow_pJP[3] + ":M" + sumrow_pJP[4] + str_enter +
                                                               ",{\"折讓-1%\",\"匯款手續費-1%\"}))";
                        wsheet_JP.Cell(i + 26, 12).FormulaA1 = "= ROUND(K" + (i + 26) + "* 0.01, 0)";
                        wsheet_JP.Cell(i + 26, 13).FormulaA1 = "= ROUND(L" + (i + 26) + "* 0.1, 0)";
                        wsheet_JP.Cell(i + 26, 14).FormulaA1 = "= L" + (i + 26) + "-M" + (i + 26);

                        wsheet_JP.Cell(i + 27, 8).Value = "2%+1%";
                        wsheet_JP.Cell(i + 27, 11).FormulaA1 = "=sum(K" + (i + 25) + ":K" + (i + 26) + ")";
                        wsheet_JP.Cell(i + 27, 12).FormulaA1 = "=sum(L" + (i + 25) + ":L" + (i + 26) + ")";
                        wsheet_JP.Cell(i + 27, 13).FormulaA1 = "=sum(M" + (i + 25) + ":M" + (i + 26) + ")";
                        wsheet_JP.Cell(i + 27, 14).FormulaA1 = "=sum(N" + (i + 25) + ":N" + (i + 26) + ")";

                        //start
                        wsheet_JP.Range("E" + (i + 30) + ":N" + (i + 51)).Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                        wsheet_JP.Cell(i + 30, 5).Value = "調整後";
                        wsheet_JP.Range("N" + (i + 30) + ":N" + (i + 44)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_JP.Cell(i + 30, 13).Value = "AU(AU-C5E)";
                        wsheet_JP.Cell(i + 30, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"AU\",C:C,\"小計\")";
                        sumrow_pJP[1] = i + 30;
                        wsheet_JP.Cell(i + 31, 13).Value = "AU-T(AU-C4A)";
                        wsheet_JP.Cell(i + 31, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"AU-T\",C:C,\"小計\")";

                        wsheet_JP.Cell(i + 32, 13).Value = "AU-TK(路竹)";
                        wsheet_JP.Cell(i + 32, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"AU-TK\",C:C,\"小計\")";

                        wsheet_JP.Cell(i + 33, 13).Value = "AU-TN(台南)";
                        wsheet_JP.Cell(i + 33, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"AU-TN\",C:C,\"小計\")";

                        wsheet_JP.Cell(i + 34, 13).Value = "TCE";
                        wsheet_JP.Cell(i + 34, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"TCE\",C:C,\"小計\")";

                        wsheet_JP.Cell(i + 35, 13).Value = "TCET-AU";
                        wsheet_JP.Cell(i + 35, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"TCET-AU\",C:C,\"小計\")";

                        for (int num = 0; num <= m; num++)
                        {
                            switch (custidname_pJP[num])
                            {
                                case "AU":
                                case "AU-T":
                                case "AU-TK":
                                case "AU-TN":
                                case "TCE":
                                case "TCET-AU":

                                case "CSOT":
                                case "WCSOT":
                                case "CPCF":
                                case "AU-L6K":
                                case "AU-TC":
                                case "AU-TY":
                                case "BOE":
                                case "CPD":
                                case "HKC":
                                case "HKC-H2":
                                case "HKC-H4":
                                case "HKC-H5":
                                case "CCPD":
                                case "CHOT":
                                case "HSD":
                                    break;
                                default:
                                    wsheet_JP.Cell(i + 36, 13).Value = custidname_pJP[num];
                                    wsheet_JP.Cell(i + 36, 14).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                                    wsheet_JP.Cell(i + 36, 14).FormulaA1 = "=-SUMIFS(N:N,A:A,\"" + custidname_pJP[num] + "\",C:C,\"小計\")";
                                    i++;
                                    break;
                            }
                            p++;
                        }
                        
                        wsheet_JP.Range("M" + (i + 36) + ":N" + (i + 39)).Style.Fill.BackgroundColor = XLColor.FromHtml("#F1DBDA");
                        wsheet_JP.Cell(i + 36, 13).Value = "折讓-2%";
                        wsheet_JP.Cell(i + 36, 14).FormulaA1 = "=-折讓_日本凸版!N" + num_dcJP[0];
                        sumrow_pJP[3] = i + 36;

                        wsheet_JP.Cell(i + 37, 13).Value = "折讓-1%";
                        wsheet_JP.Cell(i + 37, 14).FormulaA1 = "=-折讓_日本凸版!N" + num_dcJP[1];


                        wsheet_JP.Cell(i + 38, 13).Value = "匯款手續費-2%";
                        wsheet_JP.Cell(i + 38, 14).FormulaA1 = "=-手續費!H" + num_hJP[0];

                        wsheet_JP.Cell(i + 39, 13).Value = "匯款手續費-1%";
                        wsheet_JP.Cell(i + 39, 14).FormulaA1 = "=-手續費!H" + num_hJP[1];
                       

                        wsheet_JP.Range("M" + (i + 40) + ":N" + (i + 43)).Style.Fill.BackgroundColor = XLColor.FromHtml("#D8E4BC");
                        wsheet_JP.Cell(i + 40, 13).Value = "CHOT 15日~月底";
                        wsheet_JP.Cell(i + 41, 13).Value = "CPCF 21日~月底";
                        wsheet_JP.Cell(i + 42, 13).Value = "CSOT 21日~月底";
                        wsheet_JP.Cell(i + 43, 13).Value = "WCSOT 21日~月底";
                        
                        sumrow_pJP[2] = i + 43;
                        sumrow_pJP[4] = i + 43;
                        
                        switch (str_date_e2_m)
                        {
                            case "03":
                            case "06":
                            case "09":
                            case "12":
                                wsheet_JP.Cell(i + 40, 14).FormulaA1 = String.Format(
                                    @"=-SUMIFS(N:N,A:A,""CHOT"",C:C,"">={0}15"")", str_date_e2_15.Substring(0, 6));
                                wsheet_JP.Cell(i + 41, 14).FormulaA1 = String.Format(
                                    @"=-SUMIFS(N:N,A:A,""CPCF"",C:C,"">={0}21"")", str_date_e2_15.Substring(0, 6));
                                wsheet_JP.Cell(i + 42, 14).FormulaA1 = String.Format(
                                    @"=-SUMIFS(N:N,A:A,""CSOT"",C:C,"">={0}21"")", str_date_e2_15.Substring(0, 6));
                                wsheet_JP.Cell(i + 43, 14).FormulaA1 = String.Format(
                                    @"=-SUMIFS(N:N,A:A,""WCSOT"",C:C,"">={0}21"")", str_date_e2_15.Substring(0, 6));
                                break;

                            case "01":
                            case "04":
                            case "07":
                            case "10":
                                wsheet_JP.Cell(i + 40, 14).FormulaA1 = String.Format(
                                    @"=SUMIFS(N:N,A:A,""CHOT"",C:C,""<{0}01"")", str_date_e2_15.Substring(0, 6));
                                wsheet_JP.Cell(i + 41, 14).FormulaA1 = String.Format(
                                    @"=SUMIFS(N:N,A:A,""CPCF"",C:C,""<{0}01"")", str_date_e2_15.Substring(0, 6));
                                wsheet_JP.Cell(i + 42, 14).FormulaA1 = String.Format(
                                    @"=SUMIFS(N:N,A:A,""CSOT"",C:C,""<{0}01"")", str_date_e2_15.Substring(0, 6));
                                wsheet_JP.Cell(i + 43, 14).FormulaA1 = String.Format(
                                    @"=SUMIFS(N:N,A:A,""WCSOT"",C:C,""<{0}01"")", str_date_e2_15.Substring(0, 6));
                                break;

                            default:
                                break;
                        }

                        wsheet_JP.Cell(i + 44, 13).Value = "合計";
                        wsheet_JP.Cell(i + 44, 14).Style.NumberFormat.Format = "#,##0";
                        wsheet_JP.Range("M" + (i + 44) + ":N" + (i + 44)).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFBD45");
                        wsheet_JP.Cell(i + 44, 14).FormulaA1 = "=N" + sumrow_pJP[0] + "+SUM(N" + sumrow_pJP[1] + ":N" + sumrow_pJP[2] + ")";

                        i = i + 3;
                        wsheet_JP.Range("F" + (i + 44) + ":N" + (i + 47)).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFBD45");
                        wsheet_JP.Range("F" + (i + 44) + ":N" + (i + 47)).Style.Font.FontSize = 14;
                        wsheet_JP.Range("K" + (i + 45) + ":N" + (i + 47)).Style.NumberFormat.Format = "#,##0";
                        wsheet_JP.Range("G" + (i + 45) + ":G" + (i + 46)).Style.Font.FontSize = 10;

                        wsheet_JP.Range("F" + (i + 44) + ":N" + (i + 47)).Style.Border.OutsideBorder = XLBorderStyleValues.Double;

                        wsheet_JP.Range("H" + (i + 44) + ":H" + (i + 46)).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        wsheet_JP.Cell(i + 44, 11).Value = "銷貨";
                        wsheet_JP.Cell(i + 44, 12).Value = "權利金";
                        wsheet_JP.Cell(i + 44, 13).Value = "10%稅金";
                        wsheet_JP.Cell(i + 44, 14).Value = "淨額";

                        wsheet_JP.Cell(i + 45, 6).Value = "Royalty";
                        wsheet_JP.Cell(i + 45, 7).Value = "CSOT、WCSOT、CPCF、AU-L6K、AU-TC、AU-TY";
                        wsheet_JP.Cell(i + 45, 10).Value = "2%";
                        wsheet_JP.Cell(i + 45, 11).FormulaA1 = String.Format(
                            @"=SUM(SUMIFS(N:N,A:A,{{""CSOT"",""WCSOT"",""CPCF"",""AU-L6K"",""AU-TC"",""AU-TY""}},C:C,""小計"")
                              ,SUMIFS(N{0}:N{1},M{2}:M{3},{{""折讓-2%"",""匯款手續費-2%"",""CPCF*"",""CSOT*"",""WCSOT*""}}))"
                              , sumrow_pJP[3], sumrow_pJP[4], sumrow_pJP[3], sumrow_pJP[4]);

                        wsheet_JP.Cell(i + 45, 12).FormulaA1 = "= ROUND(K" + (i + 45) + "* 0.02, 0)";
                        wsheet_JP.Cell(i + 45, 13).FormulaA1 = "= ROUND(L" + (i + 45) + "* 0.1, 0)";
                        wsheet_JP.Cell(i + 45, 14).FormulaA1 = "= L" + (i + 45) + "-M" + (i + 45);

                        wsheet_JP.Cell(i + 46, 7).Value = "BOE、CPD、HKC、HKC-H2、HKC-H4、HKC-H5、CHOT、HSD、CCPD";
                        wsheet_JP.Cell(i + 46, 10).Value = "1%";
                        wsheet_JP.Cell(i + 46, 11).FormulaA1 = String.Format(
                            @"=SUM(SUMIFS(N:N,A:A,{{""BOE"",""CPD"",""HKC"",""HKC-H2"",""HKC-H4"",""HKC-H5"",""CHOT"",""HSD"",""CCPD""}},C:C,""小計"")
                              ,SUMIFS(N{0}:N{1},M{2}:M{3},{{""折讓-1%"",""匯款手續費-1%"",""CHOT*""}}))"
                              , sumrow_pJP[3], sumrow_pJP[4], sumrow_pJP[3], sumrow_pJP[4]);

                        wsheet_JP.Cell(i + 46, 12).FormulaA1 = "= ROUND(K" + (i + 46) + "* 0.01, 0)";
                        wsheet_JP.Cell(i + 46, 13).FormulaA1 = "= ROUND(L" + (i + 46) + "* 0.1, 0)";
                        wsheet_JP.Cell(i + 46, 14).FormulaA1 = "= L" + (i + 46) + "-M" + (i + 46);

                        wsheet_JP.Cell(i + 47, 8).Value = "2%+1%";
                        wsheet_JP.Cell(i + 47, 11).FormulaA1 = "=sum(K" + (i + 45) + ":K" + (i + 46) + ")";
                        wsheet_JP.Cell(i + 47, 12).FormulaA1 = "=sum(L" + (i + 45) + ":L" + (i + 46) + ")";
                        wsheet_JP.Cell(i + 47, 13).FormulaA1 = "=sum(M" + (i + 45) + ":M" + (i + 46) + ")";
                        wsheet_JP.Cell(i + 47, 14).FormulaA1 = "=sum(N" + (i + 45) + ":N" + (i + 46) + ")";
                    }
                    i++;
                }
                //worksheet.Columns().AdjustToContents();
                //worksheet2.Columns().AdjustToContents();

                wsheet_JP.Position = 1;
                wsheet_JP.Column("A").Width = 7;
                wsheet_JP.Column("B").Width = 10;
                wsheet_JP.Column("C").Width = 8;
                wsheet_JP.Column("D").Width = 4;
                wsheet_JP.Column("E").Width = 6.5;
                wsheet_JP.Column("F").Width = 7.5;
                wsheet_JP.Column("G").Width = 16;
                wsheet_JP.Column("H").Width = 11;
                wsheet_JP.Column("I").Width = 10;
                wsheet_JP.Column("J").Width = 7;

                wsheet_JP.Column("K").Width = 14;
                wsheet_JP.Column("L").Width = 14;
                wsheet_JP.Column("M").Width = 14;
                wsheet_JP.Column("N").Width = 14;

                save_as_JP = txt_path.Text.ToString().Trim() + @"\\銷貨明細(權利金)日本凸版_調整" + txt_date_e2.Text.ToString().Substring(0, 6) + "_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";

                wb_JP.SaveAs(save_as_JP);

            }
            //打开文件
            System.Diagnostics.Process.Start(save_as_JP);
            System.Diagnostics.Process.Start(save_as_MPT);
        }

        private void Btn_pre2_Click(object sender, EventArgs e)
        {
            using (XLWorkbook wb_pAll = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel))
                {
                    var ws = templateWB.Worksheet(5);

                    ws.CopyTo(wb_pAll, "權利金");
                }

                var wsheet_pAll = wb_pAll.Worksheet("權利金");

                //== 權利金報表 ================================================================
                //var worksheet = wb.Worksheets.Add("權利金_光阻");

                int rows_count_pAll = dt_pAll.Rows.Count;
                int i = 0; int y = 0;
                string cust_pAll = "";
                string sum_pAll = "";

                foreach (DataRow row in dt_pMPT.Rows)
                {
                    if (cust_pAll.ToString() != "" && row[0].ToString() != cust_pAll.ToString())
                    {
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_pAll.Cell(i + 5, 1).Value = cust_pAll;
                        wsheet_pAll.Cell(i + 5, 3).Value = "小計";
                        wsheet_pAll.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_pAll.Cell(i + 5, 8).FormulaA1 = "=sum(H" + (y + 5) + ":H" + (i + 4) + ")";
                        wsheet_pAll.Range("N" + (i + 5) + ":P" + (i + 5)).Style.NumberFormat.Format = "#,##0";
                        wsheet_pAll.Cell(i + 5, 14).FormulaA1 = "=sum(N" + (y + 5) + ":N" + (i + 4) + ")";
                        wsheet_pAll.Cell(i + 5, 15).FormulaA1 = "=sum(O" + (y + 5) + ":O" + (i + 4) + ")";
                        wsheet_pAll.Cell(i + 5, 16).FormulaA1 = "=sum(P" + (y + 5) + ":P" + (i + 4) + ")";

                        y = i + 1;
                        i++;

                    }

                    //填入excel欄位值
                    wsheet_pAll.Cell(i + 5, 1).Value = row[0]; //客戶代號
                    wsheet_pAll.Cell(i + 5, 2).Value = row[1]; //客戶簡稱
                    wsheet_pAll.Cell(i + 5, 3).Value = row[2]; //銷貨日期
                    wsheet_pAll.Cell(i + 5, 4).Style.NumberFormat.Format = "@";
                    wsheet_pAll.Cell(i + 5, 4).Value = row[3]; //單別
                    wsheet_pAll.Cell(i + 5, 5).Style.NumberFormat.Format = "@";
                    wsheet_pAll.Cell(i + 5, 5).Value = row[4]; //單號
                    wsheet_pAll.Cell(i + 5, 6).Value = row[5]; //批號
                    wsheet_pAll.Cell(i + 5, 7).Value = row[6]; //品號

                    if (row[6].ToString().Substring(0, 3) == "MPT")
                    {
                        sum_pAll += "N" + (i + 5) + "+";
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.FromHtml("#BDD7EE");
                    }

                    wsheet_pAll.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                    wsheet_pAll.Cell(i + 5, 8).Value = row[7]; //數量
                    wsheet_pAll.Cell(i + 5, 9).Style.NumberFormat.Format = "#,##0";
                    wsheet_pAll.Cell(i + 5, 9).Value = row[8]; //原幣未稅金額
                    wsheet_pAll.Cell(i + 5, 10).Style.NumberFormat.Format = "#,##0";
                    wsheet_pAll.Cell(i + 5, 10).Value = row[9]; //原幣稅額
                    wsheet_pAll.Cell(i + 5, 11).Style.NumberFormat.Format = "#,##0";
                    wsheet_pAll.Cell(i + 5, 11).Value = row[10]; //原幣合計金額
                    wsheet_pAll.Cell(i + 5, 12).Value = row[11]; //幣別
                    wsheet_pAll.Cell(i + 5, 13).Style.NumberFormat.Format = "#,##0.000";
                    wsheet_pAll.Cell(i + 5, 13).Value = row[12]; //匯率
                    wsheet_pAll.Cell(i + 5, 14).Style.NumberFormat.Format = "#,##0";
                    wsheet_pAll.Cell(i + 5, 14).Value = row[13]; //本幣未稅金額
                    wsheet_pAll.Cell(i + 5, 15).Style.NumberFormat.Format = "#,##0";
                    wsheet_pAll.Cell(i + 5, 15).Value = row[14]; //本幣稅額
                    wsheet_pAll.Cell(i + 5, 16).Style.NumberFormat.Format = "#,##0";
                    wsheet_pAll.Cell(i + 5, 16).Value = row[15]; //本幣合計金額

                    cust_pAll = row[0].ToString().Trim();

                    if ((rows_count_pAll - 1) == dt_pMPT.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        i++;
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_pAll.Range("A" + (i + 5) + ":P" + (i + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_pAll.Cell(i + 5, 1).Value = cust_pAll;
                        wsheet_pAll.Cell(i + 5, 3).Value = "小計";
                        wsheet_pAll.Cell(i + 5, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_pAll.Cell(i + 5, 8).FormulaA1 = "=sum(H" + (y + 5) + ":H" + (i + 4) + ")";
                        wsheet_pAll.Range("N" + (i + 5) + ":P" + (i + 5)).Style.NumberFormat.Format = "#,##0";
                        wsheet_pAll.Cell(i + 5, 14).FormulaA1 = "=sum(N" + (y + 5) + ":N" + (i + 4) + ")";
                        wsheet_pAll.Cell(i + 5, 15).FormulaA1 = "=sum(O" + (y + 5) + ":O" + (i + 4) + ")";
                        wsheet_pAll.Cell(i + 5, 16).FormulaA1 = "=sum(P" + (y + 5) + ":P" + (i + 4) + ")";

                        y = i + 1;

                        wsheet_pAll.Cell("B1").Value = "調整查詢";
                        wsheet_pAll.Cell("B2").Value = txt_date_s2.Text.ToString() + "~" + txt_date_e2.Text.ToString();
                        wsheet_pAll.Cell("B3").Value = createday;
                        wsheet_pAll.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_pAll.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_pAll.Cell(i + 6, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_pAll.Range("A" + (i + 6) + ":P" + (i + 6)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_pAll.Cell(i + 6, 3).Value = "總計";
                        wsheet_pAll.Cell(i + 6, 8).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_pAll.Cell(i + 6, 8).FormulaA1 = "=SUMIF(C:C,\"小計\",H:H)";
                        wsheet_pAll.Range("N" + (i + 6) + ":P" + (i + 6)).Style.NumberFormat.Format = "#,##0";
                        wsheet_pAll.Cell(i + 6, 14).FormulaA1 = "=SUMIF(C:C,\"小計\",N:N)";
                        wsheet_pAll.Cell(i + 6, 15).FormulaA1 = "=SUMIF(C:C,\"小計\",O:O)";
                        wsheet_pAll.Cell(i + 6, 16).FormulaA1 = "=SUMIF(C:C,\"小計\",P:P)";
                    }
                    i++;
                }
                //wsheet_pAll.Columns().AdjustToContents();
                //worksheet2.Columns().AdjustToContents();

                wsheet_pAll.Position = 1;
                wsheet_pAll.Column("A").Width = 7;
                wsheet_pAll.Column("B").Width = 10;
                wsheet_pAll.Column("C").Width = 8;
                wsheet_pAll.Column("D").Width = 4;
                wsheet_pAll.Column("E").Width = 6.5;
                wsheet_pAll.Column("F").Width = 7.5;
                wsheet_pAll.Column("G").Width = 16;
                wsheet_pAll.Column("H").Width = 11;
                wsheet_pAll.Column("I").Width = 10;
                wsheet_pAll.Column("J").Width = 7;


                wsheet_pAll.Column("K").Width = 14;
                wsheet_pAll.Column("L").Width = 14;
                wsheet_pAll.Column("M").Width = 14;
                wsheet_pAll.Column("N").Width = 14;

                //wsheet_pAll.Column("K").Width = 16;
                //wsheet_pAll.Column("L").Width = 16;
                //wsheet_pAll.Column("M").Width = 16;
                //wsheet_pAll.Column("N").Width = 16;

                save_as_All = txt_path.Text.ToString().Trim() + @"\\權利金_調整" + txt_date_e2.Text.ToString().Substring(0, 6) + "_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";

                wb_pAll.SaveAs(save_as_All);

                //打开文件
                System.Diagnostics.Process.Start(save_as_All);
                }
            }

        private void Btn_dc2_Click(object sender, EventArgs e)
        {
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
                        wsheet_dcAll.Cell("B3").Value = "調整查詢";
                        wsheet_dcAll.Cell("B4").Value = txt_date_s.Text.ToString() + " ~ " + txt_date_e.Text.ToString();
                        wsheet_dcAll.Cell("B5").Value = createday;
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

                save_as_dcAll = txt_path.Text.ToString().Trim() + @"\\折讓" + txt_date_e2.Text.ToString().Substring(0, 6) + "_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";

                wb_dcAll.SaveAs(save_as_dcAll);

                //打开文件
                System.Diagnostics.Process.Start(save_as_dcAll);
            }
        }
    }
}
