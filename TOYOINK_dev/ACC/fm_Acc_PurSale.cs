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
    public partial class fm_Acc_PurSale : Form
    {
        public MyClass MyCode;
        月曆 fm_月曆;
        string save_as_PurSale = "", temp_excel_PurSale;
        string createday = DateTime.Now.ToString("yyyy/MM/dd");
        int opencode = 0;

        string str_date_s, str_date_m_s, str_date_ym_s, str_date_y_s;
        string str_date_e, str_date_m_e, str_date_ym_e, str_date_y_e;

        string cond_Search;
        string defaultfilePath = "";

        DateTime date_s, date_e;
        
        DataTable dt_Search = new DataTable("dt_Search");  //查詢結果

        public fm_Acc_PurSale()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass(); ;

            //MyCode.strDbCon = MyCode.strDbConLeader;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConLeader;

            MyCode.strDbCon = MyCode.strDbConA01A;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConA01A;

            //MyCode.strDbCon = MyCode.strDbConTemp;

            temp_excel_PurSale = @"\\192.168.128.219\Conductor\Company\MIS自開發主檔\會計報表公版\進貨之銷貨_temp.xlsx";

        }


        private void fm_Acc_PurSale_Load(object sender, EventArgs e)
        {
            txt_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-01-01")).AddYears(-1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-12-31")).AddYears(-1).ToString("yyyyMMdd");
            string filder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            txt_path.Text = filder;

            cond_Search = @"and TG001 <> '233' and TH007<> '43' and MB032 <> ''";

            txterr.Text = string.Format(@"1.取[結束]抓取月份，例如：2024/06/01，將抓取[2024/06]資訊。
2.日期變更後，先前查詢資料須重新查詢，若無查詢，禁止Excel轉出。
3.Excel轉出後包含明細，程式自動開啟該報表。
4.查詢條件：
======== 品號基本檔內[主要供應商(MB032)]不可為空值 ===========
{0}", cond_Search);

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
                    var ws = templateWB.Worksheet("進貨之銷貨");

                    ws.CopyTo(wb_PurSale, "進貨之銷貨");

                }

                var wsheet_Search = wb_PurSale.Worksheet("進貨之銷貨");

                MyCode.ERP_DTInputExcel(wsheet_Search, dt_Search, 5, 1, str_date_ym_s, str_date_ym_e);

                save_as_PurSale = txt_path.Text.ToString().Trim() + "\\" + str_date_ym_e + @"_進貨之銷貨_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
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
            string sql_str_Search = String.Format(@"SELECT MB032 as 主要供應商代號
                                        ,PURMA.MA002 as 供應商簡稱,PURMA.MA085 as 供應商關係人代號
                                        ,(case when TG006='WCSOT' then 'WCSOT' else COPMA.MA002 end) as 客戶別
                                        ,COPMA.MA124 as 客戶關係人代號
                                           ,SUBSTRING(TG003,1,6) as 銷貨年月,TH004 as 品目
                                           ,sum(TH008) as 銷貨數,sum(TH037) as 本幣未稅金額,sum(LA013) as 總原價
                                           ,sum(LA017) as 材料,sum(LA018) as 人工,sum(LA019) as 製費
                                           ,TG001 as 單別,MQ002 as 單據名稱,MB005 as 會計別,MB006 as 產品別
                                           ,INVMA.MA003 as 商品
                                         FROM   S2008X64.A01A.dbo.COPTG
                                         INNER  JOIN S2008X64.A01A.dbo.COPTH ON TG001 = TH001 AND TG002 = TH002
                                         INNER  JOIN S2008X64.A01A.dbo.INVLA ON TH001=LA006 and TH002=LA007 and TH003=LA008 and TH004=LA001
                                         left join  S2008X64.A01A.dbo.INVMB ON MB001 = TH004
                                         left join S2008X64.A01A.dbo.PURMA on MB032 = PURMA.MA001
                                         left join S2008X64.A01A.dbo.INVMA ON MB007=INVMA.MA002 and MB001 = TH004
                                         left join S2008X64.A01A.dbo.CMSMQ on MQ001 = TG001
                                         left join S2008X64.A01A.dbo.COPMA on COPMA.MA001 = TG004
                                         WHERE  SUBSTRING(TG003,1,6) between '{0}' and '{1}' {2}
                                        group by (case when TG006='WCSOT' then 'WCSOT' else COPMA.MA002 end),
                                        SUBSTRING(TG003,1,6),TH004,TG001,MQ002,MB005,MB006,INVMA.MA003,MB032,PURMA.MA002,PURMA.MA085,COPMA.MA124
                                        order by SUBSTRING(TG003,1,6),MB032,COPMA.MA124,TH004"
                                        , str_date_ym_s, str_date_ym_e,cond_Search);
            MyCode.Sql_dgv(sql_str_Search, dt_Search, dgv_Search);

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
            dt_Search.Clear();
            dgv_Search.DataSource = null;

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
