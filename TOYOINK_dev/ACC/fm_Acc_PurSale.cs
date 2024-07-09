using ClosedXML.Excel;
using Myclass;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TOYOINK_dev
{
    /* 20240614 財務林姿刪 提出開發需求
     * 20240621 財務林姿刪 提出 
     * 1.銷貨單，會計01製品，會計04物料；產品別52-0JP(會計00商品)，不列入報表中
        原因是這兩種屬於光阻銷售，不需列入。
       2.銷退單，數量及金額請顯示負數
       3.佣金收入單別 2SHT、C2SH ，因為沒有成本，如電話中討論，請篩選已確認單據列入報表中
       4.三角貿易稅額調整，請參照銷貨成本分析月報5a中的明細表，只需列入項目稅額調整的數據即可
     * 20240627 財務林姿刪 提出
       1.銷退單 客戶代號空值，欄位對應錯誤，正確 TJ004
       2.佣金收入單別 left JOIN S2008X64.A01A.dbo.INVLA 
       3.須排除產品別52開頭，如 52-0JP、52-3K、52-3S、52-4CN、52-4F；MB006 not like ('52%')。
     * 
     */
    public partial class fm_Acc_PurSale : Form
    {
        public MyClass MyCode;
        月曆 fm_月曆;
        string save_as_PurSale = "", temp_excel_PurSale;
        string createday = DateTime.Now.ToString("yyyy/MM/dd");
        int opencode = 0;

        string str_date_s, str_date_m_s, str_date_ym_s, str_date_y_s;
        string str_date_e, str_date_m_e, str_date_ym_e, str_date_y_e;

        string cond_COPTH, cond_COPTJ, cond_ACTML;
        string defaultfilePath = "";

        DateTime date_s, date_e;
        
        DataTable dt_COPTH = new DataTable("dt_COPTH");  //銷貨單
        DataTable dt_COPTJ = new DataTable("dt_COPTJ");  //銷退單
        DataTable dt_ACTML = new DataTable("dt_ACTML");  //明細分類帳-稅額調整

        public fm_Acc_PurSale()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass(); ;

            //MyCode.strDbCon = MyCode.strDbConLeader;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConLeader;

            MyCode.strDbCon = MyCode.strDbConA01A;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConA01A;

            //MyCode.strDbCon = MyCode.strDbConTemp;

            temp_excel_PurSale = @"\\192.168.128.219\Conductor\Company\MIS自開發主檔\會計報表公版\進貨之銷售_temp.xlsx";

        }


        private void fm_Acc_PurSale_Load(object sender, EventArgs e)
        {
            txt_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).AddYears(-1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-12-31")).AddYears(-1).ToString("yyyyMMdd");
            string filder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            txt_path.Text = filder;

            cond_COPTH = @"and TG001 <> '233' and TH007<> '43' and MB005 not in ('01','04') and MB006 not like ('52%') and TG023 = 'Y'";
            cond_COPTJ = @"and MB005 not in ('01','04') and MB006 not like ('52%')  and TI019='Y'";
            cond_ACTML = @"ML009 like '%稅額調整%' and ML006 = '410202'";

            txterr.Text = string.Format(@"1.抓取月份，例如：2024/06/01，將抓取[2024/06]資訊，預設以【年】為單位。
2.日期變更後，先前查詢資料須重新查詢，若無查詢，禁止Excel轉出。
3.Excel轉出後包含明細，程式自動開啟該報表。
4.查詢條件：
=========== 銷貨單 ===========
{0}
=========== 銷退單 ===========
{1}
=========== 明細分類帳-稅額調整 ===========
{2}
", cond_COPTH, cond_COPTJ, cond_ACTML);

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

        private void btn_ToExcel_Click(object sender, EventArgs e)
        {
            BtnFalse();

            using (XLWorkbook wb_PurSale = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel_PurSale))
                {
                    var ws_COPTH = templateWB.Worksheet("進貨之銷貨");
                    var ws_COPTJ = templateWB.Worksheet("進貨之銷退");
                    var ws_ACTML = templateWB.Worksheet("明細分類帳");

                    ws_COPTH.CopyTo(wb_PurSale, "進貨之銷貨");
                    ws_COPTJ.CopyTo(wb_PurSale, "進貨之銷退");
                    ws_ACTML.CopyTo(wb_PurSale, "明細分類帳");

                }

                var wsheet_COPTH = wb_PurSale.Worksheet("進貨之銷貨");
                var wsheet_COPTJ = wb_PurSale.Worksheet("進貨之銷退");
                var wsheet_ACTML = wb_PurSale.Worksheet("明細分類帳");

                MyCode.ERP_DTInputExcel(wsheet_COPTH, dt_COPTH, 5, 1, str_date_ym_s, str_date_ym_e);
                MyCode.ERP_DTInputExcel(wsheet_COPTJ, dt_COPTJ, 5, 1, str_date_ym_s, str_date_ym_e);
                MyCode.ERP_DTInputExcel(wsheet_ACTML, dt_ACTML, 5, 1, str_date_ym_s, str_date_ym_e);

                save_as_PurSale = txt_path.Text.ToString().Trim() + "\\" + str_date_ym_e + @"_進貨之銷售_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
                wb_PurSale.SaveAs(save_as_PurSale);

                //打开文件
                if (opencode != 1)
                {
                    System.Diagnostics.Process.Start(save_as_PurSale);
                }
            }
            BtnTrue();
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

            //TODO:查詢進貨之銷貨
            string sql_str_COPTH = String.Format(@"SELECT MB032 as 主要供應商代號
                                        ,PURMA.MA002 as 供應商簡稱,PURMA.MA085 as 供應商關係人代號
                                        ,(case when TG006='WCSOT' then 'WCSOT' else COPMA.MA002 end) as 客戶別
                                        ,COPMA.MA124 as 客戶關係人代號
                                           ,SUBSTRING(TG003,1,4) as 銷貨年,SUBSTRING(TG003,1,6) as 銷貨年月,TH004 as 品目
                                           ,sum(TH008) as 銷貨數,sum(TH037) as 本幣未稅金額,sum(LA013) as 總原價
                                           ,sum(LA017) as 材料,sum(LA018) as 人工,sum(LA019) as 製費
                                           ,TG001 as 單別,TG002 as 單號,MQ002 as 單據名稱,MB005 as 會計別,MB006 as 產品別
                                           ,INVMA.MA003 as 商品
                                         FROM   S2008X64.A01A.dbo.COPTG
                                         INNER JOIN S2008X64.A01A.dbo.COPTH ON TG001 = TH001 AND TG002 = TH002
                                         left JOIN S2008X64.A01A.dbo.INVLA ON TH001=LA006 and TH002=LA007 and TH003=LA008 and TH004=LA001
                                         left join S2008X64.A01A.dbo.INVMB ON MB001 = TH004
                                         left join S2008X64.A01A.dbo.PURMA on MB032 = PURMA.MA001
                                         left join S2008X64.A01A.dbo.INVMA ON MB007=INVMA.MA002 and MB001 = TH004
                                         left join S2008X64.A01A.dbo.CMSMQ on MQ001 = TG001
                                         left join S2008X64.A01A.dbo.COPMA on COPMA.MA001 = TG004
                                         WHERE  SUBSTRING(TG003,1,6) between '{0}' and '{1}' {2}
                                        group by (case when TG006='WCSOT' then 'WCSOT' else COPMA.MA002 end),SUBSTRING(TG003,1,4),SUBSTRING(TG003,1,6),
                                        TH004,TG001,TG002,MQ002,MB005,MB006,INVMA.MA003,MB032,PURMA.MA002,PURMA.MA085,COPMA.MA124
                                        order by SUBSTRING(TG003,1,6),MB032,COPMA.MA124,TH004"
                                        , str_date_ym_s, str_date_ym_e,cond_COPTH);
            MyCode.Sql_dgv(sql_str_COPTH, dt_COPTH, dgv_COPTH);

            //TODO:查詢進貨之銷退
            string sql_str_COPTJ = String.Format(@"SELECT MB032 as 主要供應商代號
                                        ,PURMA.MA002 as 供應商簡稱,PURMA.MA085 as 供應商關係人代號
                                        ,(case when TI004='WCSOT' then 'WCSOT' else COPMA.MA002 end) as 客戶別
                                        ,COPMA.MA124 as 客戶關係人代號
                                            ,SUBSTRING(TI003,1,4) as 銷貨年,SUBSTRING(TI003,1,6) as 銷貨年月,TJ004 as 品目
                                            ,-sum(TJ007) as 銷貨數,-sum(TJ033) as 本幣未稅金額,-sum(LA013) as 總原價
                                            ,-sum(LA017) as 材料,-sum(LA018) as 人工,-sum(LA019) as 製費
                                            ,TI001 as 單別 ,TI002 as 單號,MQ002 as 單據名稱,MB005 as 會計別,MB006 as 產品別
                                            ,INVMA.MA003 as 商品
                                            FROM   S2008X64.A01A.dbo.COPTI
                                            INNER JOIN S2008X64.A01A.dbo.COPTJ ON TI001 = TJ001 AND TI002 = TJ002
                                            left JOIN S2008X64.A01A.dbo.INVLA ON TJ001=LA006 and TJ002=LA007 and TJ003=LA008 and TJ004=LA001
                                            left join  S2008X64.A01A.dbo.INVMB ON MB001 = TJ004
                                            left join S2008X64.A01A.dbo.PURMA on MB032 = PURMA.MA001
                                            left join S2008X64.A01A.dbo.INVMA ON MB007=INVMA.MA002 and MB001 = TJ004
                                            left join S2008X64.A01A.dbo.CMSMQ on MQ001 = TI001
                                            left join S2008X64.A01A.dbo.COPMA on COPMA.MA001 = TI004
                                         WHERE  SUBSTRING(TI003,1,6) between '{0}' and '{1}' {2}
                                        group by (case when TI004='WCSOT' then 'WCSOT' else COPMA.MA002 end),SUBSTRING(TI003,1,6),SUBSTRING(TI003,1,4),
                                        TJ004,TI001,TI002,MQ002,MB005,MB006,INVMA.MA003,MB032,PURMA.MA002,PURMA.MA085,COPMA.MA124
                                        order by SUBSTRING(TI003,1,6),MB032,COPMA.MA124,TJ004"
                                        , str_date_ym_s, str_date_ym_e, cond_COPTJ);
            MyCode.Sql_dgv(sql_str_COPTJ, dt_COPTJ, dgv_COPTJ);

            //TODO:查詢進貨之明細分類帳-稅額調整
            string sql_str_ACTML = String.Format(@"select * from (
                                      select ML006 as 科目編號 
                                        ,(select MA003 from ACTMA where MA001 = ACTML.ML006) as 科目名稱
                                        ,SUBSTRING(ML002,1,6) as 傳票年月 ,ML003+'-'+ML004+' -'+ML005 as 傳票編號
                                        ,ML009 as 摘要 ,TB012 as 備註
                                        ,(case ML007 when '1' then ML008 else 0 end) as 本幣借方金額
                                        ,(case ML007 when '-1' then ML008 else 0 end)  as 本幣貸方金額 
                                        ,(case MA007 
	                                        when '1' then ((case ML007 when '1' then ML008 else 0 end)-(case ML007 when '-1' then ML008 else 0 end)) 
	                                        when '-1' then ((case ML007 when '-1' then ML008 else 0 end)-(case ML007 when '1' then ML008 else 0 end)) else 0 end ) as 貸借金額
                                        ,case ML006
		                                    when '410202' then '稅額調整'
	                                     end as 項目
	                                     ,case ML006
		                                    when '410202' then '未稅金額' 
	                                     end as 成本
	                                    ,case ML006
		                                    when '410202' then SUBSTRING(ML009,CHARINDEX('#',ML009)+1,2) 
	                                     end as 產品別
                                        from ACTML
	                                        left JOIN ACTMA on ACTMA.MA001 = ACTML.ML006
                                            left JOIN ACTTB on ACTTB.TB001 = ACTML.ML003 and ACTTB.TB002 = ACTML.ML004 and ACTTB.TB003 = ACTML.ML005
                                        where ({2})) ACTML_ALL
                                    where 傳票年月 between '{0}' and '{1}'
                                    order by 傳票年月,科目編號"
                                        , str_date_ym_s, str_date_ym_e, cond_ACTML);
            MyCode.Sql_dgv(sql_str_ACTML, dt_ACTML, dgv_ACTML);

            BtnTrue();
        }

        private void btn_down_Click(object sender, EventArgs e)
        {
            date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s.Text = DateTime.Parse(date_s.ToString("yyyy-MM-01")).AddYears(-1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(date_e.ToString("yyyy-MM-31")).AddYears(-1).ToString("yyyyMMdd");

            DtAndDgvClear();
        }

        private void btn_up_Click(object sender, EventArgs e)
        {
            date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            txt_date_s.Text = DateTime.Parse(date_s.ToString("yyyy-MM-01")).AddYears(1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(date_e.ToString("yyyy-MM-31")).AddYears(1).ToString("yyyyMMdd");

            DtAndDgvClear();
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

        private void DtAndDgvClear()
        {
            dt_COPTH.Clear();
            dgv_COPTH.DataSource = null;

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
