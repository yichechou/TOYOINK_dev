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

namespace TOYOINK_dev
{
    public partial class fm_Trademark : Form
    {
        /* 20191104 新增.SetTabColor(XLColor.Red) 分頁顏色調整為紅色
         * 20200204 設定1月份不會自動上底色，修正底色月份，從[ for (int colornum = 1; colornum < sheetMonth; colornum++)] 改為從0 開始
         * 20200406 關係人加入'Y-PP0017'
         * 20200515 更新公版\MIS自開發主檔\會計報表公版
         * 20210303 2月份有新增銷貨單別:2SHT佣金-關係人，故要修改新增商標權程式中2個工作表勞務收入(佣金)總表、 銷貨單_勞務收入(佣金)關係人) 資料
         * 20210305 除查詢需加入2SHT外，Excel轉出仍需加入單別
         * 20210513 更改查詢條件，不指定單別，改套代號cond及判別客戶基本資料COPMA-MA124關係人代號不為空值及開頭9
         */
        public MyClass MyCode;
        string str_enter = ((char)13).ToString() + ((char)10).ToString();
        月曆 fm_月曆;

        string str_date_m_s, str_date_m_e,str_date_twy_e, str_date_y_e;
        string str_date_s, str_date_e;
        string sql_str_cond_Income,sql_str_cond_Statement,sql_str_cond_Subsidiary,sql_str_cond_Order,sql_str_cond_CustOrder;

        DataTable dt_Income = new DataTable();  //損益表
        DataTable dt_Statement = new DataTable();  //結帳單
        DataTable dt_Subsidiary = new DataTable();  //明細分類帳
        DataTable dt_OrderList = new DataTable();  //清單-銷貨單_勞務佣金
        DataTable dt_Order = new DataTable();  //銷貨單_勞務佣金
        DataTable dt_CustOrder = new DataTable();  //年度客戶別銷售
        DataTable dt_CustOrderList = new DataTable();  //年度客戶別銷售彙總

        DataSet ds = new DataSet();
        DateTime startDate, endDate;

        string save_as_Trademark = "",temp_excel;

        int totalMonth,sheetMonth;

        string[] monthlist = new string[12] ;
        string[] YearMonthlist = new string[12];

        bool err;
        string defaultfilePath = "";

        public fm_Trademark()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();
            MyCode.strDbCon = "packet size=4096;user id=pwuser;password=sqlmis003;data source=192.168.128.219;persist security info=False;initial catalog=A01A;";
            temp_excel = @"\\192.168.128.219\Company\MIS自開發主檔\會計報表公版\商標權報表_temp.xlsx";
            //temp_excel = @"D:\商標權報表_temp.xlsx";
        }

        private void txt_date_s_TextChanged(object sender, EventArgs e)
        {
            str_date_s = txt_date_s.Text.Trim();
            str_date_m_s = txt_date_s.Text.Trim().Substring(0, 6);
           
            MonthListCode();
        }

        private void txt_date_e_TextChanged(object sender, EventArgs e)
        {
            str_date_e = txt_date_e.Text.Trim();
            str_date_m_e = txt_date_e.Text.Trim().Substring(0, 6);
            
            MonthListCode();
        }

        private void Btn_date_s_Click(object sender, EventArgs e)
        {
            if (dgv_Income.DataSource != null )
            {
                DialogResult Result = MessageBox.Show("修改 單據日期 後，需重新【查詢】", "已查詢過", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                if (Result == DialogResult.OK)
                {
                    cbo_sheet.Enabled = true;
                    CleanItem();

                    this.fm_月曆 = new 月曆(this.txt_date_s, this.Btn_date_s, "單據起始日期");
                }
                else if (Result == DialogResult.Cancel)
                {
                    return;
                }
            }
            else
            {
                cbo_sheet.Enabled = true;
                CleanItem();
                this.fm_月曆 = new 月曆(this.txt_date_s, this.Btn_date_s, "單據起始日期");
            }
            
        }

        private void Btn_date_e_Click(object sender, EventArgs e)
        {
            if (dgv_Income.DataSource != null)
            {
                DialogResult Result = MessageBox.Show("修改 單據日期 後，需重新【查詢】", "已查詢過", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);

                if (Result == DialogResult.OK)
                {
                    cbo_sheet.Enabled = true;
                    CleanItem();
                    this.fm_月曆 = new 月曆(this.txt_date_e, this.Btn_date_e, "單據結束日期");
                }
                else if (Result == DialogResult.Cancel)
                {
                    return;
                }
            }
            else
            {
                cbo_sheet.Enabled = true;
                CleanItem();
                this.fm_月曆 = new 月曆(this.txt_date_e, this.Btn_date_e, "單據結束日期");
            }
            
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

        private void cbo_sheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void fm_Trademark_FormClosed(object sender, FormClosedEventArgs e)
        {
            Environment.Exit(Environment.ExitCode);
        }

        private void fm_Trademark_Load(object sender, EventArgs e)
        {
            sql_str_cond_Statement = "TA025 ='Y' and TB004 <> '9'  and MA124 <> N'' and left(MA124,1) <> '9'";
            sql_str_cond_Order = "TH020 = 'Y' and TH026 = 'Y' and TB004 <>'9' and TH027 in ('C61T','6SHT') and TH001 in ('C2SH','2SHT') and MA124 <> N'' and left(MA124,1) <> '9'";
            sql_str_cond_CustOrder = "確認碼 = 'Y' and 結帳碼 = 'Y' and MA124<> N'' and left(MA124,1) <> '9'";

            txterr.Text = string.Format(@"1.取[結束]抓取月份，例如：2021/01/29，將抓取[2021/01]資訊。
2.日期變更後，先前查詢資料須重新查詢，若無查詢，禁止Excel轉出。
3.查詢條件：
==== 結帳單_東洋集團關係人 ====
{0}
==== 銷貨單_勞務收入(佣金)關係人 ====
{1}
==== 年度客戶別銷售金額統計表 ====
{2}", sql_str_cond_Statement, sql_str_cond_Order, sql_str_cond_CustOrder);

            tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;

            //txt_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            //txt_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
            //DateTime.Parse(DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + "1").AddMonths(1).AddDays(-1).ToShortDateString();

            if (int.Parse(DateTime.Now.ToString("MM")) > 6)
            {
                txt_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).AddMonths(6).ToString("yyyyMMdd");
                txt_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
            }
            else
            {
                txt_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).ToString("yyyyMMdd");
                txt_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
            }

            string filder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            txt_path.Text = filder;

            ds.Tables.Add(dt_Income);
            ds.Tables.Add(dt_Statement);
            ds.Tables.Add(dt_Subsidiary);
            ds.Tables.Add(dt_OrderList);
            ds.Tables.Add(dt_Order);
            ds.Tables.Add(dt_CustOrder);
            ds.Tables.Add(dt_CustOrderList);


        }
       
        private void btn_search_Click(object sender, EventArgs e)
        {
            Btn_acc.Enabled = false;
            cbo_sheet.Enabled= false;


            if (MyClass.DateIntervalCheck(txt_date_s, txt_date_e) is false)
            {
                return;
            }

            MonthListCode();

            if (err == false)
            {
                //損益表
                string sql_str_Income = String.Format(
                    @"SELECT MB001 as 科目編號
                    ,(select MA003 from ACTMA where MA001 = ACTMB.MB001) as 科目名稱
                    ,Left(MB001,1) as 科目層級1,Left(MB001,2) as 科目層級2,Left(MB001,3) as 科目層級3
                    , MB002 as 會計年度, MB003 as 期別,(MB002 + MB003) as 年度,sum(MB005-MB004) as 貸借金額
                    FROM [A01A].[dbo].ACTMB
                    where (MB002 + MB003) >= '{0}' and (MB002 + MB003) <= '{1}' and MB001 like '4%'
                    group by MB001,MB002,MB003", str_date_m_s, str_date_m_e);

                 MyCode.Sql_dgv(sql_str_Income, dt_Income, dgv_Income);

                //結帳單_東洋集團關係人
                
                string sql_str_Statement = String.Format(
                    @"SELECT TA032 as 申報年月, TA004 as 客戶代號, TA008 as 客戶全名,  TB001 as 結帳單別
                    , MQ002 as 單據名稱, TA003 as 結帳日期, TB002 as 結帳單號, TB003 as 結帳序號,TB004 as 來源
                    ,(case TB013 when '' then MB012 else TB013 end)as 會計科目
                    ,TB019 as 本幣未稅金額,TB005 as 銷貨單別,TB008 as 銷貨單據日期,TB006 as 銷貨單號
                    FROM [A01A].[dbo].ACRTB
                        left JOIN ACRTA on ACRTA.TA001 = ACRTB.TB001 and  ACRTA.TA002 = ACRTB.TB002
                        left JOIN CMSMQ on CMSMQ.MQ001  = ACRTB.TB001 
                        left JOIN AJSMB on AJSMB.MB002 = ACRTB.TB001
                        left join COPMA on COPMA.MA001 = TA004
                        where {0}
                        and TA032 >= '{1}' and TA032 <= '{2}'
                        order by TA032,TB001,TA003", sql_str_cond_Statement, str_date_m_s, str_date_m_e);

                 MyCode.Sql_dgv(sql_str_Statement, dt_Statement, dgv_Statement);

                //銷貨單_勞務收入(佣金)關係人-清單
                
                string sql_str_OrderList = String.Format(
                    @"select  MB006 品種別
                    ,(case MB006
                    when '04' then '金屬塗料'
                    when '05' then '顏料'
                    when '12' then '熱溶膠'
                    when '43' then '油墨'
                    END) 類別
                    ,MA002 客戶簡稱,TG004 客戶代號
                    , TH004 品號
                    from COPTH
                        left JOIN COPTG on COPTG.TG001=COPTH.TH001 and COPTG.TG002=COPTH.TH002
                        left JOIN COPMA on COPMA.MA001=COPTG.TG004 
                        left JOIN INVMB on INVMB.MB001=COPTH.TH004
                        left JOIN ACRTB on ACRTB.TB001 = COPTH.TH027 and ACRTB.TB002 = COPTH.TH028 and ACRTB.TB003 = COPTH.TH029
                    where {0}
                    and left(TG003,6) >= '{1}' and left(TG003,6) <= '{2}' 
                    group by TG004,MA002,MB006,TH004
                    order by MB006", sql_str_cond_Order, str_date_y_e + "01", str_date_m_e);

                MyCode.Sql_dt(sql_str_OrderList, dt_OrderList);

                //明細分類帳
                string sql_str_Subsidiary = String.Format(
                    @"select ML006 as 科目編號 
                    ,(select MA003 from ACTMA where MA001 = ACTML.ML001) as 科目名稱
                    ,ML002 as 傳票日期 ,ML003+'-'+ML004+' -'+ML005 as 傳票編號
                    ,ML009 as 摘要 ,TB012 as 備註
                    ,(case ML007 when '1' then ML008 else 0 end) as 本幣借方金額
                    ,(case ML007 when '-1' then ML008 else 0 end)  as 本幣貸方金額 
                    from ACTML
                      left JOIN ACTTB on ACTTB.TB001 = ACTML.ML003 and ACTTB.TB002 = ACTML.ML004 and ACTTB.TB003 = ACTML.ML005
                    where ML006 = '420101' and left(ML002,6) >='{0}' and left(ML002,6) <= '{1}'
                    order by ML002 ", str_date_y_e+"01", str_date_m_e);

                MyCode.Sql_dgv(sql_str_Subsidiary, dt_Subsidiary, dgv_Subsidiary);

                //銷貨單_勞務收入(佣金)關係人
                string sql_str_Order = String.Format(
                    @"select TG003 單據日期, left(TG003,6) 單據年月,TG004 客戶代號,MA002 客戶簡稱
                    ,TH001 銷貨單別,MQ002 單據名稱, TH002 銷貨單號, TB004 來源, MB006 品種別
                    , TH004 品號, TH027 結帳單別, TH028 結帳單號, TH029 結帳序號, TH037 本幣未稅金額
                    from COPTH
                      left JOIN COPTG on COPTG.TG001=COPTH.TH001 and COPTG.TG002=COPTH.TH002
                      left JOIN COPMA on COPMA.MA001=COPTG.TG004 
                      left JOIN INVMB on INVMB.MB001=COPTH.TH004
                      left JOIN CMSMQ on CMSMQ.MQ001  = COPTH.TH001 
                      left JOIN ACRTB on ACRTB.TB001 = COPTH.TH027 and ACRTB.TB002 = COPTH.TH028 and ACRTB.TB003 = COPTH.TH029
                    where {0}
                    and left(TG003,6) >= '{1}' and left(TG003,6) <= '{2}' 
                    order by MB006,TG003", sql_str_cond_Order, str_date_y_e + "01", str_date_m_e);

                MyCode.Sql_dgv(sql_str_Order, dt_Order, dgv_Order);

                //關聯方銷貨彙總表 年度客戶別銷售金額統計表 清單
                
                string sql_str_CustOrderList = String.Format(
                    @"select 客戶代號,客戶簡稱,品種別,sum(數量) as 全期銷貨量,sum(本幣未稅金額) as 全期金額
                    from (select   TG004 as 客戶代號 , MA002 as 客戶簡稱
                    , SUBSTRING(TG003,1,4) 年份 , SUBSTRING(TG003,5,2) 月份 , SUBSTRING(TG003,1,6) 年月
                    , MB006 as 品種別, TH004 as 品號, TH005 as 品名 , TH009 as 單位, sum(TH008) as 數量, sum(TH037) as 本幣未稅金額
                    , TH020 as 確認碼 , TH026 as 結帳碼
                    from COPTH
                    left join COPTG on COPTG.TG001 = COPTH.TH001 and  COPTG.TG002 = COPTH.TH002 
                    left join COPMA on COPMA.MA001 = COPTG.TG004 
                    left join INVMB on INVMB.MB001=COPTH.TH004
                    group by COPTG.TG004,MA002,MB006,TH004,TH005,TH009,SUBSTRING(TG003,5,2) , SUBSTRING(TG003,1,4), SUBSTRING(TG003,1,6),TH020,TH026
                     UNION 
                    select   TI004 as 客戶代號 , MA002 as 客戶簡稱
                    ,SUBSTRING(TI003,1,4) 年份,SUBSTRING(TI003,5,2) 月份,SUBSTRING(TI003,1,6) 年月
                    , MB006 as 品種別, TJ004 as 品號
                    , TJ005 as 品名, TJ008 as 單位, sum(-TJ007) as 數量 , sum(-TJ033) as 本幣未稅金額
                    , TJ021 as 確認碼, TJ024 as 結帳碼
                    from COPTJ
                    left join COPTI on COPTI.TI001 = COPTJ.TJ001 and  COPTI.TI002 = COPTJ.TJ002 
                    left join COPMA on COPMA.MA001 = COPTI.TI004 
                    left join INVMB on INVMB.MB001=COPTJ.TJ004
                    group by COPTI.TI004,MA002,MB006,TJ004,TJ005,TJ008,SUBSTRING(TI003,5,2),SUBSTRING(TI003,1,4),SUBSTRING(TI003,1,6),TJ021,TJ024 ) a
                    left join COPMA on COPMA.MA001 = 客戶代號
                    where {0}
                    and 年月 >= '{1}' and 年月 <= '{2}'
                    group by 客戶代號,客戶簡稱,品種別", sql_str_cond_CustOrder, str_date_y_e + "01", str_date_m_e);

                MyCode.Sql_dt(sql_str_CustOrderList, dt_CustOrderList);

                //年度客戶別銷售金額統計表
                string sql_str_CustOrder = String.Format(
                    @"select 客戶代號,客戶簡稱,品種別,品號,品名,單位
                    ,SUM(CASE WHEN 月份 =01THEN 數量 ELSE 0 END) '01銷貨量'
                    ,SUM(CASE WHEN 月份 =01THEN 本幣未稅金額 ELSE 0 END) '01金額'
                    ,SUM(CASE WHEN 月份 =02THEN 數量 ELSE 0 END) '02銷貨量'
                    ,SUM(CASE WHEN 月份 =02THEN 本幣未稅金額 ELSE 0 END) '02金額'
                    ,SUM(CASE WHEN 月份 =03THEN 數量 ELSE 0 END) '03銷貨量'
                    ,SUM(CASE WHEN 月份 =03THEN 本幣未稅金額 ELSE 0 END) '03金額'
                    ,SUM(CASE WHEN 月份 =04THEN 數量 ELSE 0 END) '04銷貨量'
                    ,SUM(CASE WHEN 月份 =04THEN 本幣未稅金額 ELSE 0 END) '04金額'
                    ,SUM(CASE WHEN 月份 =05THEN 數量 ELSE 0 END) '05銷貨量'
                    ,SUM(CASE WHEN 月份 =05THEN 本幣未稅金額 ELSE 0 END) '05金額'
                    ,SUM(CASE WHEN 月份 =06THEN 數量 ELSE 0 END) '06銷貨量'
                    ,SUM(CASE WHEN 月份 =06THEN 本幣未稅金額 ELSE 0 END) '06金額'
                    ,SUM(CASE WHEN 月份 =07THEN 數量 ELSE 0 END) '07銷貨量'
                    ,SUM(CASE WHEN 月份 =07THEN 本幣未稅金額 ELSE 0 END) '07金額'
                    ,SUM(CASE WHEN 月份 =08THEN 數量 ELSE 0 END) '08銷貨量'
                    ,SUM(CASE WHEN 月份 =08THEN 本幣未稅金額 ELSE 0 END) '08金額'
                    ,SUM(CASE WHEN 月份 =09THEN 數量 ELSE 0 END) '09銷貨量'
                    ,SUM(CASE WHEN 月份 =09THEN 本幣未稅金額 ELSE 0 END) '09金額'
                    ,SUM(CASE WHEN 月份 =10THEN 數量 ELSE 0 END) '10銷貨量'
                    ,SUM(CASE WHEN 月份 =10THEN 本幣未稅金額 ELSE 0 END) '10金額'
                    ,SUM(CASE WHEN 月份 =11THEN 數量 ELSE 0 END) '11銷貨量'
                    ,SUM(CASE WHEN 月份 =11THEN 本幣未稅金額 ELSE 0 END) '11金額'
                    ,SUM(CASE WHEN 月份 =12THEN 數量 ELSE 0 END) '12銷貨量'
                    ,SUM(CASE WHEN 月份 =12THEN 本幣未稅金額 ELSE 0 END) '12金額'
                    ,sum(數量) as 全期銷貨量,sum(本幣未稅金額) as 全期金額
                    from (select   TG004 as 客戶代號 , MA002 as 客戶簡稱
                    , SUBSTRING(TG003,1,4) 年份 , SUBSTRING(TG003,5,2) 月份 , SUBSTRING(TG003,1,6) 年月
                    , MB006 as 品種別, TH004 as 品號, TH005 as 品名 , TH009 as 單位, sum(TH008) as 數量, sum(TH037) as 本幣未稅金額
                    , TH020 as 確認碼 , TH026 as 結帳碼
                    from COPTH
                    left join COPTG on COPTG.TG001 = COPTH.TH001 and  COPTG.TG002 = COPTH.TH002 
                    left join COPMA on COPMA.MA001 = COPTG.TG004 
                    left join INVMB on INVMB.MB001=COPTH.TH004
                    group by COPTG.TG004,MA002,MB006,TH004,TH005,TH009,SUBSTRING(TG003,5,2) , SUBSTRING(TG003,1,4), SUBSTRING(TG003,1,6),TH020,TH026
                     UNION 
                    select   TI004 as 客戶代號 , MA002 as 客戶簡稱
                    ,SUBSTRING(TI003,1,4) 年份,SUBSTRING(TI003,5,2) 月份,SUBSTRING(TI003,1,6) 年月
                    , MB006 as 品種別, TJ004 as 品號
                    , TJ005 as 品名, TJ008 as 單位, sum(-TJ007) as 數量 , sum(-TJ033) as 本幣未稅金額
                    , TJ021 as 確認碼, TJ024 as 結帳碼
                    from COPTJ
                    left join COPTI on COPTI.TI001 = COPTJ.TJ001 and  COPTI.TI002 = COPTJ.TJ002 
                    left join COPMA on COPMA.MA001 = COPTI.TI004 
                    left join INVMB on INVMB.MB001=COPTJ.TJ004
                    group by COPTI.TI004,MA002,MB006,TJ004,TJ005,TJ008,SUBSTRING(TI003,5,2),SUBSTRING(TI003,1,4),SUBSTRING(TI003,1,6),TJ021,TJ024 ) a
                    left join COPMA on COPMA.MA001 = 客戶代號
                    where {0}
                    and 年月 >= '{1}' and 年月 <= '{2}'
                    group by 客戶代號,客戶簡稱,品種別,品號,品名,單位", sql_str_cond_CustOrder, str_date_y_e + "01", str_date_m_e);

                MyCode.Sql_dgv(sql_str_CustOrder, dt_CustOrder, dgv_CustOrder);

                Btn_acc.Enabled = true;
            }
        }

        private void Btn_acc_Click(object sender, EventArgs e)
        {
            string colorM;
            string[] colorMlist = new string[12];
            string[] colorMnum = new string[12];
            int sheetnum = int.Parse(cbo_sheet.Text.ToString().Trim());

            using (XLWorkbook wb_Trademark = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel))
                {
                    string name = sheetnum.ToString() + "商標權";
                    var ws = templateWB.Worksheet(name);
                    var ws2 = templateWB.Worksheet("損益表");
                    var ws3 = templateWB.Worksheet("結帳單_東洋集團關係人");
                    var ws4 = templateWB.Worksheet("勞務收入(佣金)總表");
                    var ws5 = templateWB.Worksheet("明細分類帳_勞務收入(佣金)關係人");
                    var ws6 = templateWB.Worksheet("銷貨單_勞務收入(佣金)關係人");
                    var ws7 = templateWB.Worksheet("關聯方銷貨彙總表");
                    var ws8 = templateWB.Worksheet("年度客戶別銷售金額統計表");

                    //ADDFUN 20191104 新增 .SetTabColor(XLColor.Red) 分頁顏色調整為紅色
                    ws.CopyTo(wb_Trademark, "商標權").SetTabColor(XLColor.Red);
                    ws2.CopyTo(wb_Trademark, "損益表");
                    ws3.CopyTo(wb_Trademark, "結帳單_東洋集團關係人");
                    ws4.CopyTo(wb_Trademark, "勞務收入(佣金)總表").SetTabColor(XLColor.Red);
                    ws5.CopyTo(wb_Trademark, "明細分類帳_勞務收入(佣金)關係人");
                    ws6.CopyTo(wb_Trademark, "銷貨單_勞務收入(佣金)關係人");
                    ws7.CopyTo(wb_Trademark, "關聯方銷貨彙總表").SetTabColor(XLColor.Red);
                    ws8.CopyTo(wb_Trademark, "年度客戶別銷售金額統計表").SetTabColor(XLColor.Red);
                }

                var wsheet_Trademark = wb_Trademark.Worksheet("商標權");
                var wsheet_Income = wb_Trademark.Worksheet("損益表");
                var wsheet_Statement = wb_Trademark.Worksheet("結帳單_東洋集團關係人");
                var wsheet_OrderList = wb_Trademark.Worksheet("勞務收入(佣金)總表");
                var wsheet_Subsidiary = wb_Trademark.Worksheet("明細分類帳_勞務收入(佣金)關係人");
                var wsheet_Order = wb_Trademark.Worksheet("銷貨單_勞務收入(佣金)關係人");
                var wsheet_CustOrderList = wb_Trademark.Worksheet("關聯方銷貨彙總表");
                var wsheet_CustOrder = wb_Trademark.Worksheet("年度客戶別銷售金額統計表");

                //=== 商標權 3.52==========================================
                wsheet_Trademark.Cell(1, 3).Value = "會計年度:" + str_date_twy_e ; //會計年度
                wsheet_Trademark.Cell(2, 2).Value = "製表日期:" + DateTime.Now.ToString("yyyy/MM/dd"); //會計年度
                wsheet_Trademark.Cell(2, 3).Value = "月份區間:" + str_date_m_s + "~" + str_date_m_e; //查詢月份區間

                wsheet_Trademark.Cell(3, 4).Value = monthlist[0]; //年月
                colorMlist[0] = monthlist[0] ;
                colorMnum[0] = "4";

                for (int monthnum = 1; monthnum < sheetMonth; monthnum++)
                {
                    wsheet_Trademark.Cell(3, 4 + monthnum).Value = monthlist[monthnum]; //年月
                    colorMlist[monthnum] = monthlist[monthnum];
                    colorMnum[monthnum] = (4 + monthnum).ToString();
                }

                if (chk_colorM.Checked)
                {
                    colorM = cbo_colorM.Text.ToString().Trim();

                    for (int colornum = 0; colornum < sheetMonth; colornum++)
                    {
                        if (colorM == colorMlist[colornum])
                        {
                            wsheet_Trademark.Cell(3, colorMnum[colornum]).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wsheet_Trademark.Cell(53, colorMnum[colornum]).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }
                    }
                }

                ////== 損益表 ================================================================
                int rows_count_Income = dt_Income.Rows.Count;
                int i = 0;

                wsheet_Income.Cell(2, 2).Value = str_date_m_s + "~" + str_date_m_e; //查詢月份區間
                wsheet_Income.Cell(3, 2).Style.NumberFormat.Format = "@";
                wsheet_Income.Cell(3, 2).Value =  DateTime.Now.ToString("yyyy/MM/dd"); //製表日期
                foreach (DataRow row in dt_Income.Rows)
                {
                    wsheet_Income.Cell(i + 5, 1).Value = row[0]; //科目編號
                    wsheet_Income.Cell(i + 5, 2).Value = row[1]; //科目名稱
                    wsheet_Income.Cell(i + 5, 3).Value = row[2]; //科目層級1
                    wsheet_Income.Cell(i + 5, 4).Value = row[3]; //科目層級2
                    wsheet_Income.Cell(i + 5, 5).Value = row[4]; //科目層級3
                    wsheet_Income.Cell(i + 5, 6).Value = row[5]; //會計年度
                    wsheet_Income.Cell(i + 5, 7).Style.NumberFormat.Format = "@";
                    wsheet_Income.Cell(i + 5, 7).Value = row[6]; //期別
                    wsheet_Income.Cell(i + 5, 8).Value = row[7]; //年度
                    wsheet_Income.Cell(i + 5, 9).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                    wsheet_Income.Cell(i + 5, 9).Value = row[8]; //借貸金額

                    i++;
                }

                ////== 結帳單 ================================================================
                int rows_count_Statement = dt_Statement.Rows.Count;
                int j = 0;
                string Snewj = "";
                string Soldj = "";

                wsheet_Statement.Cell(2, 2).Value = str_date_m_s + "~" + str_date_m_e; //查詢月份區間
                wsheet_Statement.Cell(3, 2).Style.NumberFormat.Format = "@";
                wsheet_Statement.Cell(3, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期

                foreach (DataRow row in dt_Statement.Rows)
                {
                    Snewj = row[0].ToString();
                    if (Soldj != Snewj && j != 0)
                    {
                        wsheet_Statement.Range("A" + (j + 5) + ":N" + (j + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_Statement.Range("A" + (j + 5) + ":N" + (j + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_Statement.Range("A" + (j + 5) + ":N" + (j + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;

                        wsheet_Statement.Cell(j + 5, 2).Value = Soldj; //申報年月
                        wsheet_Statement.Cell(j + 5, 3).Value = "小計";
                        wsheet_Statement.Cell(j + 5, 11).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_Statement.Cell(j + 5, 11).FormulaA1 = "=SUMIF(A:A,\"" + Soldj +"\",K:K)";
                       
                        j++;
                    }

                    Soldj = row[0].ToString();

                    wsheet_Statement.Cell(j + 5, 1).Value = row[0]; //申報年月
                    wsheet_Statement.Cell(j + 5, 2).Value = row[1]; //客戶代號
                    wsheet_Statement.Cell(j + 5, 3).Value = row[2]; //客戶全名
                    wsheet_Statement.Cell(j + 5, 4).Style.NumberFormat.Format = "@";
                    wsheet_Statement.Cell(j + 5, 4).Value = row[3]; //結帳單別
                    wsheet_Statement.Cell(j + 5, 5).Value = row[4]; //單據名稱
                    wsheet_Statement.Cell(j + 5, 6).Value = row[5]; //結帳日期
                    wsheet_Statement.Cell(j + 5, 7).Style.NumberFormat.Format = "@";
                    wsheet_Statement.Cell(j + 5, 7).Value = row[6]; //結帳單號
                    wsheet_Statement.Cell(j + 5, 8).Style.NumberFormat.Format = "@";
                    wsheet_Statement.Cell(j + 5, 8).Value = row[7]; //結帳序號
                    wsheet_Statement.Cell(j + 5, 9).Value = row[8]; //來源
                    wsheet_Statement.Cell(j + 5, 10).Value = row[9]; //會計科目
                    wsheet_Statement.Cell(j + 5, 11).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                    wsheet_Statement.Cell(j + 5, 11).Value = row[10]; //本幣未稅金額
                    wsheet_Statement.Cell(j + 5, 12).Style.NumberFormat.Format = "@";
                    wsheet_Statement.Cell(j + 5, 12).Value = row[11]; //銷貨單別
                    wsheet_Statement.Cell(j + 5, 13).Value = row[12]; //銷貨單據日期
                    wsheet_Statement.Cell(j + 5, 14).Style.NumberFormat.Format = "@";
                    wsheet_Statement.Cell(j + 5, 14).Value = row[13]; //銷貨單號

                    if ((rows_count_Statement - 1) == dt_Statement.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        j++;
                        wsheet_Statement.Range("A" + (j + 5) + ":N" + (j + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_Statement.Range("A" + (j + 5) + ":N" + (j + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_Statement.Range("A" + (j + 5) + ":N" + (j + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;

                        wsheet_Statement.Cell(j + 5, 2).Value = Soldj; //申報年月
                        wsheet_Statement.Cell(j + 5, 3).Value = "小計";
                        wsheet_Statement.Cell(j + 5, 11).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_Statement.Cell(j + 5, 11).FormulaA1 = "=SUMIF(A:A,\"" + Soldj + "\",K:K)";

                        j++;
                        wsheet_Statement.Range("A" + (j + 5) + ":N" + (j + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_Statement.Range("A" + (j + 5) + ":N" + (j + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_Statement.Range("A" + (j + 5) + ":N" + (j + 5)).Style.Fill.BackgroundColor = XLColor.Honeydew;

                        wsheet_Statement.Cell(j + 5, 3).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_Statement.Cell(j + 5, 3).Value = "總計";
                        wsheet_Statement.Cell(j + 5, 11).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_Statement.Cell(j + 5, 11).FormulaA1 = "=SUMIF(C:C,\"小計\",K:K)";
                    }
                    j++;
                }

                //=== 勞務收入(佣金)總表 -清單 ==========================================
                int rows_count_OrderList = dt_OrderList.Rows.Count;
                int n = 0;

                wsheet_OrderList.Cell(1, 2).Value = "會計年度:" + str_date_twy_e; //會計年度

                for (int monthnum = 0; monthnum < 12; monthnum++)
                {
                    wsheet_OrderList.Cell(2, 4 + monthnum).Value = YearMonthlist[monthnum]; //年月
                }

                foreach (DataRow row in dt_OrderList.Rows)
                {
                    wsheet_OrderList.Cell(n + 3, 1).Style.NumberFormat.Format = "@";
                    wsheet_OrderList.Cell(n + 3, 1).Value = row[0]; //品種別
                    wsheet_OrderList.Cell(n + 3, 2).Value = row[1]; //類別
                    wsheet_OrderList.Cell(n + 3, 3).Value = row[2]; //客戶簡稱
                    wsheet_OrderList.Range("D" + (n + 3) + ":P" + (n + 3)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                    
                    wsheet_OrderList.Cell(n + 3, 4).FormulaA1 = "=SUMIFS('銷貨單_勞務收入(佣金)關係人'!N:N,'銷貨單_勞務收入(佣金)關係人'!$B:$B,D$2,'銷貨單_勞務收入(佣金)關係人'!$C:$C,\"" + row[3] + "\",'銷貨單_勞務收入(佣金)關係人'!$I:$I,$A" + (n + 3) + ")"; //01月
                    wsheet_OrderList.Cell(n + 3, 5).FormulaA1 = "=SUMIFS('銷貨單_勞務收入(佣金)關係人'!N:N,'銷貨單_勞務收入(佣金)關係人'!$B:$B,E$2,'銷貨單_勞務收入(佣金)關係人'!$C:$C,\"" + row[3] + "\",'銷貨單_勞務收入(佣金)關係人'!$I:$I,$A" + (n + 3) + ")"; //02月
                    wsheet_OrderList.Cell(n + 3, 6).FormulaA1 = "=SUMIFS('銷貨單_勞務收入(佣金)關係人'!N:N,'銷貨單_勞務收入(佣金)關係人'!$B:$B,F$2,'銷貨單_勞務收入(佣金)關係人'!$C:$C,\"" + row[3] + "\",'銷貨單_勞務收入(佣金)關係人'!$I:$I,$A" + (n + 3) + ")"; //03月
                    wsheet_OrderList.Cell(n + 3, 7).FormulaA1 = "=SUMIFS('銷貨單_勞務收入(佣金)關係人'!N:N,'銷貨單_勞務收入(佣金)關係人'!$B:$B,G$2,'銷貨單_勞務收入(佣金)關係人'!$C:$C,\"" + row[3] + "\",'銷貨單_勞務收入(佣金)關係人'!$I:$I,$A" + (n + 3) + ")"; //04月
                    wsheet_OrderList.Cell(n + 3, 8).FormulaA1 = "=SUMIFS('銷貨單_勞務收入(佣金)關係人'!N:N,'銷貨單_勞務收入(佣金)關係人'!$B:$B,H$2,'銷貨單_勞務收入(佣金)關係人'!$C:$C,\"" + row[3] + "\",'銷貨單_勞務收入(佣金)關係人'!$I:$I,$A" + (n + 3) + ")"; //05月
                    wsheet_OrderList.Cell(n + 3, 9).FormulaA1 = "=SUMIFS('銷貨單_勞務收入(佣金)關係人'!N:N,'銷貨單_勞務收入(佣金)關係人'!$B:$B,I$2,'銷貨單_勞務收入(佣金)關係人'!$C:$C,\"" + row[3] + "\",'銷貨單_勞務收入(佣金)關係人'!$I:$I,$A" + (n + 3) + ")"; //06月
                    wsheet_OrderList.Cell(n + 3, 10).FormulaA1 = "=SUMIFS('銷貨單_勞務收入(佣金)關係人'!N:N,'銷貨單_勞務收入(佣金)關係人'!$B:$B,J$2,'銷貨單_勞務收入(佣金)關係人'!$C:$C,\"" + row[3] + "\",'銷貨單_勞務收入(佣金)關係人'!$I:$I,$A" + (n + 3) + ")"; //07月
                    wsheet_OrderList.Cell(n + 3, 11).FormulaA1 = "=SUMIFS('銷貨單_勞務收入(佣金)關係人'!N:N,'銷貨單_勞務收入(佣金)關係人'!$B:$B,K$2,'銷貨單_勞務收入(佣金)關係人'!$C:$C,\"" + row[3] + "\",'銷貨單_勞務收入(佣金)關係人'!$I:$I,$A" + (n + 3) + ")"; //08月
                    wsheet_OrderList.Cell(n + 3, 12).FormulaA1 = "=SUMIFS('銷貨單_勞務收入(佣金)關係人'!N:N,'銷貨單_勞務收入(佣金)關係人'!$B:$B,L$2,'銷貨單_勞務收入(佣金)關係人'!$C:$C,\"" + row[3] + "\",'銷貨單_勞務收入(佣金)關係人'!$I:$I,$A" + (n + 3) + ")"; //09月
                    wsheet_OrderList.Cell(n + 3, 13).FormulaA1 = "=SUMIFS('銷貨單_勞務收入(佣金)關係人'!N:N,'銷貨單_勞務收入(佣金)關係人'!$B:$B,M$2,'銷貨單_勞務收入(佣金)關係人'!$C:$C,\"" + row[3] + "\",'銷貨單_勞務收入(佣金)關係人'!$I:$I,$A" + (n + 3) + ")"; //10月
                    wsheet_OrderList.Cell(n + 3, 14).FormulaA1 = "=SUMIFS('銷貨單_勞務收入(佣金)關係人'!N:N,'銷貨單_勞務收入(佣金)關係人'!$B:$B,N$2,'銷貨單_勞務收入(佣金)關係人'!$C:$C,\"" + row[3] + "\",'銷貨單_勞務收入(佣金)關係人'!$I:$I,$A" + (n + 3) + ")"; //11月
                    wsheet_OrderList.Cell(n + 3, 15).FormulaA1 = "=SUMIFS('銷貨單_勞務收入(佣金)關係人'!N:N,'銷貨單_勞務收入(佣金)關係人'!$B:$B,O$2,'銷貨單_勞務收入(佣金)關係人'!$C:$C,\"" + row[3] + "\",'銷貨單_勞務收入(佣金)關係人'!$I:$I,$A" + (n + 3) + ")"; //12月
                    wsheet_OrderList.Cell(n + 3, 16).FormulaA1 = "=SUM(D" + (n + 3) + ":O" + (n + 3) + ")"; //合計

                    if ((rows_count_OrderList - 1) == dt_OrderList.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        n++;
                        //wsheet_OrderList.Range("A" + (n + 3) + ":P" + (n + 3)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        //wsheet_OrderList.Range("A" + (n + 3) + ":P" + (n + 3)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_OrderList.Range("A" + (n + 3) + ":P" + (n + 3)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;

                        wsheet_OrderList.Cell(n + 3, 1).Value = "合計";
                        wsheet_OrderList.Range("D" + (n + 3) + ":P" + (n + 3)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_OrderList.Cell(n + 3, 4).FormulaA1 = "=SUM(D3:D" + (n + 2) + ")"; //01月
                        wsheet_OrderList.Cell(n + 3, 5).FormulaA1 = "=SUM(E3:E" + (n + 2) + ")"; //02月
                        wsheet_OrderList.Cell(n + 3, 6).FormulaA1 = "=SUM(F3:F" + (n + 2) + ")"; //03月
                        wsheet_OrderList.Cell(n + 3, 7).FormulaA1 = "=SUM(G3:G" + (n + 2) + ")"; //04月
                        wsheet_OrderList.Cell(n + 3, 8).FormulaA1 = "=SUM(H3:H" + (n + 2) + ")"; //05月
                        wsheet_OrderList.Cell(n + 3, 9).FormulaA1 = "=SUM(I3:I" + (n + 2) + ")"; //06月
                        wsheet_OrderList.Cell(n + 3, 10).FormulaA1 = "=SUM(J3:J" + (n + 2) + ")"; //07月
                        wsheet_OrderList.Cell(n + 3, 11).FormulaA1 = "=SUM(K3:K" + (n + 2) + ")"; //08月
                        wsheet_OrderList.Cell(n + 3, 12).FormulaA1 = "=SUM(L3:L" + (n + 2) + ")"; //09月
                        wsheet_OrderList.Cell(n + 3, 13).FormulaA1 = "=SUM(M3:M" + (n + 2) + ")"; //10月
                        wsheet_OrderList.Cell(n + 3, 14).FormulaA1 = "=SUM(N3:N" + (n + 2) + ")"; //11月
                        wsheet_OrderList.Cell(n + 3, 15).FormulaA1 = "=SUM(O3:O" + (n + 2) + ")"; //12月
                        wsheet_OrderList.Cell(n + 3, 16).FormulaA1 = "=SUM(P3:P" + (n + 2) + ")"; //合計

                        wsheet_OrderList.Range("A3:P" + (n + 3)).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        wsheet_OrderList.Range("A3:P" + (n + 3)).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }
                    n++;
                }

                ////== 明細分類帳 ================================================================
                int rows_count_Subsidiary = dt_Subsidiary.Rows.Count;
                int k = 0;

                wsheet_Subsidiary.Cell(2, 2).Value = str_date_y_e + "01" + "~" + str_date_m_e; //查詢月份區間
                wsheet_Subsidiary.Cell(3, 2).Style.NumberFormat.Format = "@";
                wsheet_Subsidiary.Cell(3, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期
                foreach (DataRow row in dt_Subsidiary.Rows)
                {
                    wsheet_Subsidiary.Cell(k + 5, 1).Value = row[0]; //科目編號
                    wsheet_Subsidiary.Cell(k + 5, 2).Value = row[1]; //科目名稱
                    wsheet_Subsidiary.Cell(k + 5, 3).Value = row[2]; //傳票日期
                    wsheet_Subsidiary.Cell(k + 5, 4).Value = row[3]; //傳票編號
                    wsheet_Subsidiary.Cell(k + 5, 5).Style.NumberFormat.Format = "@";
                    wsheet_Subsidiary.Cell(k + 5, 5).Value = row[4]; //摘要
                    wsheet_Subsidiary.Cell(k + 5, 6).Style.NumberFormat.Format = "@";
                    wsheet_Subsidiary.Cell(k + 5, 6).Value = row[5]; //備註
                    wsheet_Subsidiary.Cell(k + 5, 7).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                    wsheet_Subsidiary.Cell(k + 5, 7).Value = row[6]; //本幣借方金額
                    wsheet_Subsidiary.Cell(k + 5, 8).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                    wsheet_Subsidiary.Cell(k + 5, 8).Value = row[7]; //本幣貸方金額

                    if ((rows_count_Subsidiary - 1) == dt_Subsidiary.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        k++;
                        wsheet_Subsidiary.Range("A" + (k + 5) + ":H" + (k + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_Subsidiary.Range("A" + (k + 5) + ":H" + (k + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_Subsidiary.Range("A" + (k + 5) + ":H" + (k + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;

                        wsheet_Subsidiary.Cell(k + 5, 3).Value = "合計";
                        wsheet_Subsidiary.Cell(k + 5, 8).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_Subsidiary.Cell(k + 5, 8).FormulaA1 = "=SUMIF(A:A,\"420101\",H:H)";
                    }
                    k++;
                }

                ////== 銷貨單_勞務收入(佣金)關係人 ================================================================
                int rows_count_Order = dt_Order.Rows.Count;
                int m = 0;

                wsheet_Order.Cell(2, 2).Value = str_date_y_e + "01" + "~" + str_date_m_e; //查詢月份區間
                wsheet_Order.Cell(3, 2).Style.NumberFormat.Format = "@";
                wsheet_Order.Cell(3, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期
                foreach (DataRow row in dt_Order.Rows)
                {
                    wsheet_Order.Cell(m + 5, 1).Value = row[0]; //單據日期
                    wsheet_Order.Cell(m + 5, 2).Value = row[1]; //單據年月
                    wsheet_Order.Cell(m + 5, 3).Value = row[2]; //客戶代號
                    wsheet_Order.Cell(m + 5, 4).Value = row[3]; //客戶簡稱
                    wsheet_Order.Cell(m + 5, 5).Style.NumberFormat.Format = "@";
                    wsheet_Order.Cell(m + 5, 5).Value = row[4]; //銷貨單別
                    wsheet_Order.Cell(m + 5, 6).Value = row[5]; //單據名稱
                    wsheet_Order.Cell(m + 5, 7).Style.NumberFormat.Format = "@";
                    wsheet_Order.Cell(m + 5, 7).Value = row[6]; //銷貨單號
                    wsheet_Order.Cell(m + 5, 8).Value = row[7]; //來源
                    wsheet_Order.Cell(m + 5, 9).Style.NumberFormat.Format = "@";
                    wsheet_Order.Cell(m + 5, 9).Value = row[8]; //品種別
                    wsheet_Order.Cell(m + 5, 10).Value = row[9]; //品號
                    wsheet_Order.Cell(m + 5, 11).Style.NumberFormat.Format = "@";
                    wsheet_Order.Cell(m + 5, 11).Value = row[10]; //結帳單別
                    wsheet_Order.Cell(m + 5, 12).Style.NumberFormat.Format = "@";
                    wsheet_Order.Cell(m + 5, 12).Value = row[11]; //結帳單號
                    wsheet_Order.Cell(m + 5, 13).Style.NumberFormat.Format = "@";
                    wsheet_Order.Cell(m + 5, 13).Value = row[12]; //結帳序號
                    wsheet_Order.Cell(m + 5, 14).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                    wsheet_Order.Cell(m + 5, 14).Value = row[13]; //本幣未稅金額

                    if ((rows_count_Order - 1) == dt_Order.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        m++;
                        wsheet_Order.Range("A" + (m + 5) + ":N" + (m + 5)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_Order.Range("A" + (m + 5) + ":N" + (m + 5)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_Order.Range("A" + (m + 5) + ":N" + (m + 5)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;

                        wsheet_Order.Cell(m + 5, 3).Value = "合計";
                        wsheet_Order.Cell(m + 5, 14).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        //20210305 加入2SHT 單別
                        wsheet_Order.Cell(m + 5, 14).FormulaA1 = "=SUM(SUMIFS(N:N,E:E,{\"C2SH\",\"2SHT\"}))";
                    }

                    m++;
                }

                //var wsheet_CustOrderList = wb_Trademark.Worksheet("關聯方銷貨彙總表");
                //var wsheet_CustOrder = wb_Trademark.Worksheet("年度客戶別銷售金額統計表");
                //====== 關聯方銷貨彙總表 =====================
                int rows_count_CustOrderList = dt_CustOrderList.Rows.Count;
                int p = 0;

                wsheet_CustOrderList.Cell(4, 2).Value = str_date_y_e + "01" + "~" + str_date_m_e; //查詢月份區間
                wsheet_CustOrderList.Cell(5, 2).Style.NumberFormat.Format = "@";
                wsheet_CustOrderList.Cell(5, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期
                foreach (DataRow row in dt_CustOrderList.Rows)
                {
                    wsheet_CustOrderList.Cell(p + 7, 1).Value = row[0]; //客戶代號
                    wsheet_CustOrderList.Cell(p + 7, 2).Value = row[1]; //客戶簡稱
                    wsheet_CustOrderList.Cell(p + 7, 3).Style.NumberFormat.Format = "@";
                    wsheet_CustOrderList.Cell(p + 7, 3).Value = row[2]; //品種別
                    wsheet_CustOrderList.Cell(p + 7, 4).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                    wsheet_CustOrderList.Cell(p + 7, 4).FormulaA1 = "=SUMIFS(年度客戶別銷售金額統計表!AE:AE,年度客戶別銷售金額統計表!$A:$A,$A" + (p + 7) + ",年度客戶別銷售金額統計表!$C:$C,$C" + (p + 7) + ")";
                    wsheet_CustOrderList.Cell(p + 7, 5).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                    wsheet_CustOrderList.Cell(p + 7, 5).FormulaA1 = "=SUMIFS(年度客戶別銷售金額統計表!AF:AF,年度客戶別銷售金額統計表!$A:$A,$A" + (p + 7) + ",年度客戶別銷售金額統計表!$C:$C,$C" + (p + 7) + ")";

                    if ((rows_count_CustOrderList - 1) == dt_CustOrderList.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        p++;
                        wsheet_CustOrderList.Range("A" + (p + 7) + ":E" + (p + 7)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_CustOrderList.Range("A" + (p + 7) + ":E" + (p + 7)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_CustOrderList.Range("A" + (p + 7) + ":E" + (p + 7)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;

                        wsheet_CustOrderList.Cell(p + 7, 1).Value = "合計";
                        wsheet_CustOrderList.Cell(p + 7, 4).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_CustOrderList.Cell(p + 7, 4).FormulaA1 = "=SUM(D7:D" + (p + 6) + ")";
                        wsheet_CustOrderList.Cell(p + 7, 5).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_CustOrderList.Cell(p + 7, 5).FormulaA1 = "=SUM(E7:E" + (p + 6) + ")";
                    }

                    p++;
                }

                //====== 年度客戶別銷售金額統計表 =====================
                int rows_count_CustOrder = dt_CustOrder.Rows.Count;
                int q = 0;
                string Snewq = "";
                string Soldq = "";

                wsheet_CustOrder.Cell(3, 2).Value = str_date_y_e + "01" + "~" + str_date_m_e; //查詢月份區間
                wsheet_CustOrder.Cell(4, 2).Style.NumberFormat.Format = "@";
                wsheet_CustOrder.Cell(4, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期
                wsheet_CustOrder.Cell(4, 16).Value = "統計年度:"+ str_date_twy_e; //統計年度
                foreach (DataRow row in dt_CustOrder.Rows)
                {
                    Snewq = row[0].ToString();
                    if (Soldq != Snewq && q != 0)
                    {
                        wsheet_CustOrder.Range("A" + (q + 6) + ":AF" + (q + 6)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_CustOrder.Range("A" + (q + 6) + ":AF" + (q + 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_CustOrder.Range("A" + (q + 6) + ":AF" + (q + 6)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;

                        wsheet_CustOrder.Cell(q + 6, 6).Value = "小計";
                        wsheet_CustOrder.Range("G" + (q + 6) + ":AF" + (q + 6)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_CustOrder.Cell(q + 6, 7).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",G:G)"; //01銷貨量
                        wsheet_CustOrder.Cell(q + 6, 8).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",H:H)"; //01金額
                        wsheet_CustOrder.Cell(q + 6, 9).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",I:I)"; //02銷貨量
                        wsheet_CustOrder.Cell(q + 6, 10).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",J:J)"; //02金額
                        wsheet_CustOrder.Cell(q + 6, 11).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",K:K)"; //03銷貨量
                        wsheet_CustOrder.Cell(q + 6, 12).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",L:L)"; //03金額
                        wsheet_CustOrder.Cell(q + 6, 13).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",M:M)"; //04銷貨量
                        wsheet_CustOrder.Cell(q + 6, 14).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",N:N)"; //04金額
                        wsheet_CustOrder.Cell(q + 6, 15).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",O:O)"; //05銷貨量
                        wsheet_CustOrder.Cell(q + 6, 16).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",P:P)"; //05金額
                        wsheet_CustOrder.Cell(q + 6, 17).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",Q:Q)"; //06銷貨量
                        wsheet_CustOrder.Cell(q + 6, 18).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",R:R)"; //06金額
                        wsheet_CustOrder.Cell(q + 6, 19).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",S:S)"; //07銷貨量
                        wsheet_CustOrder.Cell(q + 6, 20).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",T:T)"; //07金額
                        wsheet_CustOrder.Cell(q + 6, 21).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",U:U)"; //08銷貨量
                        wsheet_CustOrder.Cell(q + 6, 22).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",V:V)"; //08金額
                        wsheet_CustOrder.Cell(q + 6, 23).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",W:W)"; //09銷貨量
                        wsheet_CustOrder.Cell(q + 6, 24).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",X:X)"; //09金額
                        wsheet_CustOrder.Cell(q + 6, 25).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",Y:Y)"; //10銷貨量
                        wsheet_CustOrder.Cell(q + 6, 26).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",Z:Z)"; //10金額
                        wsheet_CustOrder.Cell(q + 6, 27).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",AA:AA)"; //11銷貨量
                        wsheet_CustOrder.Cell(q + 6, 28).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",AB:AB)"; //11金額
                        wsheet_CustOrder.Cell(q + 6, 29).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",AC:AC)"; //12銷貨量
                        wsheet_CustOrder.Cell(q + 6, 30).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",AD:AD)"; //12金額
                        wsheet_CustOrder.Cell(q + 6, 31).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",AE:AE)"; //全期銷貨量
                        wsheet_CustOrder.Cell(q + 6, 32).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",AF:AF)"; //全期金額

                        q++;
                    }

                    Soldq = row[0].ToString();

                    wsheet_CustOrder.Cell(q + 6, 1).Value = row[0]; //客戶代號
                    wsheet_CustOrder.Cell(q + 6, 2).Value = row[1]; //客戶簡稱
                    wsheet_CustOrder.Cell(q + 6, 3).Style.NumberFormat.Format = "@";
                    wsheet_CustOrder.Cell(q + 6, 3).Value = row[2]; //品種別
                    wsheet_CustOrder.Cell(q + 6, 4).Value = row[3]; //品號
                    wsheet_CustOrder.Cell(q + 6, 5).Value = row[4]; //品名
                    wsheet_CustOrder.Cell(q + 6, 6).Value = row[5]; //單位
                    wsheet_CustOrder.Range("G" + (q + 6) + ":AF" + (q + 6)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                    wsheet_CustOrder.Cell(q + 6, 7).Value = row[6]; //01銷貨量
                    wsheet_CustOrder.Cell(q + 6, 8).Value = row[7]; //01金額
                    wsheet_CustOrder.Cell(q + 6, 9).Value = row[8]; //02銷貨量
                    wsheet_CustOrder.Cell(q + 6, 10).Value = row[9]; //02金額
                    wsheet_CustOrder.Cell(q + 6, 11).Value = row[10]; //03銷貨量
                    wsheet_CustOrder.Cell(q + 6, 12).Value = row[11]; //03金額
                    wsheet_CustOrder.Cell(q + 6, 13).Value = row[12]; //04銷貨量
                    wsheet_CustOrder.Cell(q + 6, 14).Value = row[13]; //04金額
                    wsheet_CustOrder.Cell(q + 6, 15).Value = row[14]; //05銷貨量
                    wsheet_CustOrder.Cell(q + 6, 16).Value = row[15]; //05金額
                    wsheet_CustOrder.Cell(q + 6, 17).Value = row[16]; //06銷貨量
                    wsheet_CustOrder.Cell(q + 6, 18).Value = row[17]; //06金額
                    wsheet_CustOrder.Cell(q + 6, 19).Value = row[18]; //07銷貨量
                    wsheet_CustOrder.Cell(q + 6, 20).Value = row[19]; //07金額
                    wsheet_CustOrder.Cell(q + 6, 21).Value = row[20]; //08銷貨量
                    wsheet_CustOrder.Cell(q + 6, 22).Value = row[21]; //08金額
                    wsheet_CustOrder.Cell(q + 6, 23).Value = row[22]; //09銷貨量
                    wsheet_CustOrder.Cell(q + 6, 24).Value = row[23]; //09金額
                    wsheet_CustOrder.Cell(q + 6, 25).Value = row[24]; //10銷貨量
                    wsheet_CustOrder.Cell(q + 6, 26).Value = row[25]; //10金額
                    wsheet_CustOrder.Cell(q + 6, 27).Value = row[26]; //11銷貨量
                    wsheet_CustOrder.Cell(q + 6, 28).Value = row[27]; //11金額
                    wsheet_CustOrder.Cell(q + 6, 29).Value = row[28]; //12銷貨量
                    wsheet_CustOrder.Cell(q + 6, 30).Value = row[29]; //12金額
                    wsheet_CustOrder.Cell(q + 6, 31).Value = row[30]; //全期銷貨量
                    wsheet_CustOrder.Cell(q + 6, 32).Value = row[31]; //全期金額

                    if ((rows_count_CustOrder - 1) == dt_CustOrder.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        q++;
                        wsheet_CustOrder.Range("A" + (q + 6) + ":AF" + (q + 6)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_CustOrder.Range("A" + (q + 6) + ":AF" + (q + 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_CustOrder.Range("A" + (q + 6) + ":AF" + (q + 6)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;

                        wsheet_CustOrder.Cell(q + 6, 6).Value = "小計";
                        wsheet_CustOrder.Range("G" + (q + 6) + ":AF" + (q + 6)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_CustOrder.Cell(q + 6, 7).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",G:G)"; //01銷貨量
                        wsheet_CustOrder.Cell(q + 6, 8).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",H:H)"; //01金額
                        wsheet_CustOrder.Cell(q + 6, 9).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",I:I)"; //02銷貨量
                        wsheet_CustOrder.Cell(q + 6, 10).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",J:J)"; //02金額
                        wsheet_CustOrder.Cell(q + 6, 11).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",K:K)"; //03銷貨量
                        wsheet_CustOrder.Cell(q + 6, 12).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",L:L)"; //03金額
                        wsheet_CustOrder.Cell(q + 6, 13).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",M:M)"; //04銷貨量
                        wsheet_CustOrder.Cell(q + 6, 14).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",N:N)"; //04金額
                        wsheet_CustOrder.Cell(q + 6, 15).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",O:O)"; //05銷貨量
                        wsheet_CustOrder.Cell(q + 6, 16).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",P:P)"; //05金額
                        wsheet_CustOrder.Cell(q + 6, 17).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",Q:Q)"; //06銷貨量
                        wsheet_CustOrder.Cell(q + 6, 18).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",R:R)"; //06金額
                        wsheet_CustOrder.Cell(q + 6, 19).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",S:S)"; //07銷貨量
                        wsheet_CustOrder.Cell(q + 6, 20).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",T:T)"; //07金額
                        wsheet_CustOrder.Cell(q + 6, 21).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",U:U)"; //08銷貨量
                        wsheet_CustOrder.Cell(q + 6, 22).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",V:V)"; //08金額
                        wsheet_CustOrder.Cell(q + 6, 23).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",W:W)"; //09銷貨量
                        wsheet_CustOrder.Cell(q + 6, 24).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",X:X)"; //09金額
                        wsheet_CustOrder.Cell(q + 6, 25).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",Y:Y)"; //10銷貨量
                        wsheet_CustOrder.Cell(q + 6, 26).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",Z:Z)"; //10金額
                        wsheet_CustOrder.Cell(q + 6, 27).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",AA:AA)"; //11銷貨量
                        wsheet_CustOrder.Cell(q + 6, 28).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",AB:AB)"; //11金額
                        wsheet_CustOrder.Cell(q + 6, 29).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",AC:AC)"; //12銷貨量
                        wsheet_CustOrder.Cell(q + 6, 30).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",AD:AD)"; //12金額
                        wsheet_CustOrder.Cell(q + 6, 31).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",AE:AE)"; //全期銷貨量
                        wsheet_CustOrder.Cell(q + 6, 32).FormulaA1 = "=SUMIF(A:A,\"" + Soldq + "\",AF:AF)"; //全期金額

                        q++;
                        wsheet_CustOrder.Range("A" + (q + 6) + ":AF" + (q + 6)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_CustOrder.Range("A" + (q + 6) + ":AF" + (q + 6)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_CustOrder.Range("A" + (q + 6) + ":AF" + (q + 6)).Style.Fill.BackgroundColor = XLColor.Honeydew;

                        wsheet_CustOrder.Cell(q + 6, 6).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_CustOrder.Cell(q + 6, 6).Value = "總計";
                        wsheet_CustOrder.Range("G" + (q + 6) + ":AF" + (q + 6)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet_CustOrder.Cell(q + 6, 7).FormulaA1 = "=SUMIF(F:F,\"小計\",G:G)"; //01銷貨量
                        wsheet_CustOrder.Cell(q + 6, 8).FormulaA1 = "=SUMIF(F:F,\"小計\",H:H)"; //01金額
                        wsheet_CustOrder.Cell(q + 6, 9).FormulaA1 = "=SUMIF(F:F,\"小計\",I:I)"; //02銷貨量
                        wsheet_CustOrder.Cell(q + 6, 10).FormulaA1 = "=SUMIF(F:F,\"小計\",J:J)"; //02金額
                        wsheet_CustOrder.Cell(q + 6, 11).FormulaA1 = "=SUMIF(F:F,\"小計\",K:K)"; //03銷貨量
                        wsheet_CustOrder.Cell(q + 6, 12).FormulaA1 = "=SUMIF(F:F,\"小計\",L:L)"; //03金額
                        wsheet_CustOrder.Cell(q + 6, 13).FormulaA1 = "=SUMIF(F:F,\"小計\",M:M)"; //04銷貨量
                        wsheet_CustOrder.Cell(q + 6, 14).FormulaA1 = "=SUMIF(F:F,\"小計\",N:N)"; //04金額
                        wsheet_CustOrder.Cell(q + 6, 15).FormulaA1 = "=SUMIF(F:F,\"小計\",O:O)"; //05銷貨量
                        wsheet_CustOrder.Cell(q + 6, 16).FormulaA1 = "=SUMIF(F:F,\"小計\",P:P)"; //05金額
                        wsheet_CustOrder.Cell(q + 6, 17).FormulaA1 = "=SUMIF(F:F,\"小計\",Q:Q)"; //06銷貨量
                        wsheet_CustOrder.Cell(q + 6, 18).FormulaA1 = "=SUMIF(F:F,\"小計\",R:R)"; //06金額
                        wsheet_CustOrder.Cell(q + 6, 19).FormulaA1 = "=SUMIF(F:F,\"小計\",S:S)"; //07銷貨量
                        wsheet_CustOrder.Cell(q + 6, 20).FormulaA1 = "=SUMIF(F:F,\"小計\",T:T)"; //07金額
                        wsheet_CustOrder.Cell(q + 6, 21).FormulaA1 = "=SUMIF(F:F,\"小計\",U:U)"; //08銷貨量
                        wsheet_CustOrder.Cell(q + 6, 22).FormulaA1 = "=SUMIF(F:F,\"小計\",V:V)"; //08金額
                        wsheet_CustOrder.Cell(q + 6, 23).FormulaA1 = "=SUMIF(F:F,\"小計\",W:W)"; //09銷貨量
                        wsheet_CustOrder.Cell(q + 6, 24).FormulaA1 = "=SUMIF(F:F,\"小計\",X:X)"; //09金額
                        wsheet_CustOrder.Cell(q + 6, 25).FormulaA1 = "=SUMIF(F:F,\"小計\",Y:Y)"; //10銷貨量
                        wsheet_CustOrder.Cell(q + 6, 26).FormulaA1 = "=SUMIF(F:F,\"小計\",Z:Z)"; //10金額
                        wsheet_CustOrder.Cell(q + 6, 27).FormulaA1 = "=SUMIF(F:F,\"小計\",AA:AA)"; //11銷貨量
                        wsheet_CustOrder.Cell(q + 6, 28).FormulaA1 = "=SUMIF(F:F,\"小計\",AB:AB)"; //11金額
                        wsheet_CustOrder.Cell(q + 6, 29).FormulaA1 = "=SUMIF(F:F,\"小計\",AC:AC)"; //12銷貨量
                        wsheet_CustOrder.Cell(q + 6, 30).FormulaA1 = "=SUMIF(F:F,\"小計\",AD:AD)"; //12金額
                        wsheet_CustOrder.Cell(q + 6, 31).FormulaA1 = "=SUMIF(F:F,\"小計\",AE:AE)"; //全期銷貨量
                        wsheet_CustOrder.Cell(q + 6, 32).FormulaA1 = "=SUMIF(F:F,\"小計\",AF:AF)"; //全期金額
                    }
                    q++;
                }

                //wsheet_dcAll.Position = 1;

                save_as_Trademark = txt_path.Text.ToString().Trim() + @"\\商標權" + txt_date_s.Text.ToString().Substring(0, 6) +"-" + txt_date_e.Text.ToString().Substring(0, 6) + "_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
                wb_Trademark.SaveAs(save_as_Trademark);

                //打开文件
                System.Diagnostics.Process.Start(save_as_Trademark);
            }
        }

        private void CleanItem()
        {
            Btn_acc.Enabled = false;
            dgv_Income.DataSource = null;
            dgv_Statement.DataSource = null;
            dgv_Order.DataSource = null;

            ds.Clear();

            cbo_colorM.Items.Clear();
            cbo_colorM.Text = "";
        }

        private void MonthListCode()
        {
            startDate = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            endDate = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            str_date_twy_e = MyClass.ToTaiwanDateYM(endDate).Substring(0,3);
            str_date_y_e = txt_date_e.Text.Trim().Substring(0, 4);

            totalMonth = (endDate.Year - startDate.Year) * 12 + (endDate.Month - startDate.Month) + 1;
            sheetMonth = int.Parse(cbo_sheet.Text.ToString().Trim());
            txt_Mnum.Text = totalMonth.ToString();

            if (totalMonth > sheetMonth)
            {
                MessageBox.Show("請修改【日期區間】或【所需月份】", "【日期區間】大於【所需月份】", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Btn_acc.Enabled = false;
                cbo_sheet.Enabled = true;
                err = true;
                return;
            }
            else
            {
                //底色月份清單
                monthlist[0] = startDate.ToString("yyyyMM");
                for (int monthnum = 1; monthnum < sheetMonth; monthnum++)
                {
                    monthlist[monthnum] = startDate.AddMonths(monthnum).ToString("yyyyMM");
                }
                
                err = false;
                
            }

            cbo_colorM.Items.Clear();
            //選取到當月
            for (int monthnum = 0; monthnum < sheetMonth; monthnum++)
            {
                cbo_colorM.Items.Add(monthlist[monthnum]);
                if (monthlist[monthnum] == str_date_m_e)
                {
                    cbo_colorM.SelectedIndex = monthnum;
                }
            }

            //月份
            YearMonthlist = MyClass.YearMonthList(txt_date_e, 12);
        }
    }
}
