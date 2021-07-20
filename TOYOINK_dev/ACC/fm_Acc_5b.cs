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
    public partial class fm_Acc_5b : Form
    {
        //20200604 開發完成 建立者：周怡甄 需求者：鄭玉菁
        //20210118 財務林姿刪提出，去除347單別
        //20210720 財務林姿刪提出，去除進退貨條件式【(case when PURTG.TG005 = N'TVS' AND PURTG.TG007 = N'JPY' then 0 else 1 end=1)】
        public MyClass MyCode;
        月曆 fm_月曆;

        DataTable dt_5b = new DataTable("5b彙總");  //5b彙總
        DataTable dt_PURTH = new DataTable("進貨單明細");  //進貨單明細
        DataTable dt_PURTJ = new DataTable("退貨單明細");  //退貨單明細
        DataTable dt_IPS_Now = new DataTable("當月在途倉明細");  //在途倉明細(當月)加
        DataTable dt_IPS_Last = new DataTable("去年在途倉明細");  //在途倉明細(去年)減 
        DataTable dt_IPS_Now_Sum = new DataTable("當月在途彙總");  //在途倉明細(當月)加 彙總
        DataTable dt_IPS_Last_Sum = new DataTable("去年在途彙總");  //在途倉明細(去年)減 彙總

        string createday = DateTime.Now.ToString("yyyy/MM/dd");

        //string str_date_s;
        //string str_date_e, str_date_e_ym;

        string str_date_s, str_date_s_ym;
        string str_date_e, str_date_e_ym, str_date_e_y,str_date_e_lasty;

        string defaultfilePath = "", temp_excel_5b, save_as_5b = "", save_as_PUR, save_as_IPS;
        string path, fileNameWithExtension;
        bool err;

        string cond_5b, cond_PURTH, cond_PURTJ, cond_IPS_Now, cond_IPS_Last;

        DateTime date_s, date_e;
        public fm_Acc_5b()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();
            MyCode.strDbCon = "packet size=4096;user id=pwuser;password=sqlmis003;data source=192.168.128.219;persist security info=False;initial catalog=A01A;";
            //MyCode.strDbCon = "packet size=4096;user id=yj.chou;password=yjchou3369;data source=192.168.128.219;persist security info=False;initial catalog=Leader;";
            temp_excel_5b = @"\\192.168.128.219\Company\MIS自開發主檔\會計報表公版\關聯方進貨淨額明細5b_temp.xlsx";
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
            //累計
            dt_5b.Clear();   //5b彙總表
            dt_PURTH.Clear();   //進貨明細表
            dt_PURTJ.Clear();   //退貨明細表
            dt_IPS_Now.Clear();   //在途倉明細(當月)加
            dt_IPS_Last.Clear();   //在途倉明細(去年)減
            dt_IPS_Now_Sum.Clear();   //在途倉明細(當月)加
            dt_IPS_Last_Sum.Clear();   //在途倉明細(去年)減

            dgv_5b.DataSource = null;
            dgv_PURTH.DataSource = null;
            dgv_IPS_Now.DataSource = null;
            dgv_IPS_Last.DataSource = null;
            dgv_IPS_Now_Sum.DataSource = null;
            dgv_IPS_Last_Sum.DataSource = null;

            BtnFalse();
        }


        private void tab_IPS_Now_Sum_Click(object sender, EventArgs e)
        {

        }

        private void BtnFalse()
        {
            btn_5b.Enabled = false;
            btn_PUR.Enabled = false;
            btn_IPS.Enabled = false;
        }
        private void BtnTrue()
        {
            btn_5b.Enabled = true;
            btn_PUR.Enabled = true;
            btn_IPS.Enabled = true;
        }
        private void txt_date_s_TextChanged(object sender, EventArgs e)
        {
            BtnFalse();

            if (dgv_5b.DataSource != null)
            {
                DtAndDgvClear();
            }

        }

        private void txt_date_e_TextChanged(object sender, EventArgs e)
        {
            BtnFalse();

            if (dgv_5b.DataSource != null)
            {
                DtAndDgvClear();
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

        private void btn_search_Click(object sender, EventArgs e)
        {
            if (MyClass.DateIntervalCheck(txt_date_s, txt_date_e) is false)
            {
                return;
            }

            DtAndDgvClear();

            str_date_s = txt_date_s.Text.Trim();
            str_date_s_ym = txt_date_s.Text.Trim().Substring(0, 6);
            str_date_e = txt_date_e.Text.Trim();
            str_date_e_ym = txt_date_e.Text.Trim().Substring(0, 6);
            str_date_e_y = txt_date_e.Text.Trim().Substring(0, 4);

            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            str_date_e_lasty = DateTime.Parse(date_e.ToString("yyyy-MM-01")).AddYears(-1).ToString("yyyy");

            if (err == false)
            {
                //5b 明細彙總表
                string sql_str_5b = String.Format(
                    @"select 產品別,關係人代號,關係人簡稱,sum(原物料進貨淨額) as 原物料進貨淨額,sum(商品進貨淨額) as 商品進貨淨額 from 
                        (
                        --=====20200508 進貨單明細(縮減)================
                        SELECT left(INVMB.MB006,2) as 產品別
                            ,PURMA.MA085 as 關係人代號
                            ,(case PURMA.MA085 when '82000' then '東洋科美' else PURMA.MA002 End) as 關係人簡稱
                            , ROUND((case when TH001 in('345','347','349') then PURTH.TH047 else 0 End),0) as 原物料進貨淨額
                            , ROUND((case when TH001 in('340','C340') then PURTH.TH047 else 0 End),0) as 商品進貨淨額
                             FROM PURTH as PURTH  
                             Left JOIN PURTG as PURTG On PURTH.TH001=PURTG.TG001 and PURTH.TH002=PURTG.TG002 
                             Left JOIN INVMB as INVMB On PURTH.TH004=INVMB.MB001 
                             Left JOIN PURMA as PURMA On PURTG.TG005=PURMA.MA001
                         WHERE {0}
                             And left(PURTG.TG014,6) between '{1}' and '{2}'
                         union all
                        --=====20210317 退貨單明細(縮減)================
                        SELECT left(INVMB.MB006,2) as 產品別
                            ,PURMA.MA085 as 關係人代號
                            ,(case PURMA.MA085 when '82000' then '東洋科美' else PURMA.MA002 End) as 關係人簡稱
                            , ROUND((case when TJ001 in('355','359') then PURTJ.TJ032 else 0 End),0)*-1 as 原物料進貨淨額
                            , ROUND((case when TJ001 in('350','C350') then PURTJ.TJ032 else 0 End),0)*-1 as 商品進貨淨額
                             FROM PURTJ as PURTJ  
                             Left JOIN PURTI as PURTI On PURTJ.TJ001=PURTI.TI001 and PURTJ.TJ002=PURTI.TI002 
                             Left JOIN INVMB as INVMB On PURTJ.TJ004=INVMB.MB001 
                             Left JOIN PURMA as PURMA On PURTI.TI004=PURMA.MA001
                         WHERE {5}
                            And left(PURTI.TI014,6) between '{1}' and '{2}'
                         union all
                         --======20200587 在途倉明細(精簡)=====================
                        select 
                             left(MB006,2) as 產品別
                            ,PURMA.MA085 as 關係人代號
                            ,(case PURMA.MA085 when '82000' then '東洋科美' else PURMA.MA002 End) as 關係人簡稱
                            ,ROUND((case when MB005 in ('03','04') then (TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)) else 0 End),0) as 原物料進貨淨額
                            ,ROUND((case when MB005 in ('03','04') then 0 else (TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)) End),0) as 商品進貨淨額
                         from IPSTF
                            left join IPSTE on TE001 = TF001
                            left join PURMA on MA001 = TE003
                            left join PURTD on TD001 = TF002 and TD002 = TF003 and TD003 = TF004
                            left join PURTC on TC001 = TD001 and TC002 = TD002 
                            left join INVMB on MB001 = TD004
                        where left(TF015,6) = '{3}'
                        union all
                        --======20200507 去年底在途倉明細(縮減)=====================
                        select 
                             left(MB006,2) as 產品別
                            ,PURMA.MA085 as 關係人代號
                            ,(case PURMA.MA085 when '82000' then '東洋科美' else PURMA.MA002 End) as 關係人簡稱
                            ,ROUND((case when MB005 in ('03','04') then -(TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)) else 0 End),0) as 原物料進貨淨額
                            ,ROUND((case when MB005 in ('03','04') then 0 else -(TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)) End),0) as 商品進貨淨額
                         from IPSTF
                            left join IPSTE on TE001 = TF001
                            left join PURMA on MA001 = TE003
                            left join PURTD on TD001 = TF002 and TD002 = TF003 and TD003 = TF004
                            left join PURTC on TC001 = TD001 and TC002 = TD002 
                            left join INVMB on MB001 = TD004
                        where left(TF015,6) = '{4}'
                        ) b
                        group by 關係人代號,關係人簡稱,產品別"
                         , cond_PURTH, str_date_e_ym.Substring(0,4) + "01", str_date_e_ym
                         , str_date_e_ym
                         , str_date_e_lasty + "12", cond_PURTJ);

                MyCode.Sql_dgv(sql_str_5b, dt_5b, dgv_5b);

                //PURTH 進貨單明細表
                string sql_str_PURTH = String.Format(
                     @"SELECT left(INVMB.MB006,2) as 產品別,PURMA.MA085 as 關係人代號,PURTG.TG005 as 廠商代號,PURMA.MA002 as 廠商簡稱
                    , (case when TH001 in('345','347','349') then '原料' when TH001 in('340','C340') then '商品' else '' End) as 進貨類別
                    ,PURTG.TG014 as 單據日期,PURTH.TH001 as 進貨單別,PURTH.TH002 as 進貨單號,PURTH.TH004 as 品號,PURTH.TH005 as 品名
                    ,PURTH.TH008 as 單位,PURTG.TG007 as 幣別,PURTG.TG008 as 匯率,PURTH.TH007 as 進貨數量
                    ,PURTH.TH019 as 原幣進貨金額,ROUND(PURTH.TH047,0) as 本幣未稅金額
                     FROM PURTH as PURTH  
                     Left JOIN PURTG as PURTG On PURTH.TH001=PURTG.TG001 and PURTH.TH002=PURTG.TG002 
                     Left JOIN INVMB as INVMB On PURTH.TH004=INVMB.MB001 
                     Left JOIN PURMA as PURMA On PURTG.TG005=PURMA.MA001
                     WHERE {0}
                     And left(PURTG.TG014,6) between '{1}' and '{2}'
                     ORDER BY left(INVMB.MB006,2) asc,PURMA.MA085 asc,PURTG.TG014 asc,PURTH.TH001 asc,PURTH.TH002 asc"
                    , cond_PURTH, str_date_e_ym.Substring(0, 4) + "01", str_date_e_ym);

                MyCode.Sql_dgv(sql_str_PURTH, dt_PURTH, dgv_PURTH);

                //PURTJ	 退貨單明細表
                string sql_str_PURTJ = String.Format(
                     @"SELECT left(INVMB.MB006,2) as 產品別,PURMA.MA085 as 關係人代號,PURTI.TI004 as 廠商代號,PURMA.MA002 as 廠商簡稱
                    ,(case when TJ001 in('355','359') then '原料' when TJ001 in('350','C350') then '商品' else '' End) as 退貨類別
                    ,PURTI.TI014 as 單據日期,PURTJ.TJ001 as 退貨單別,PURTJ.TJ002 as 退貨單號,PURTJ.TJ004 as 品號,PURTJ.TJ005 as 品名
                    ,PURTJ.TJ007 as 單位,PURTI.TI006 as 幣別,PURTI.TI007 as 匯率,PURTJ.TJ009*-1 as 退貨數量
                    ,PURTJ.TJ030*-1 as 原幣退貨金額,ROUND(PURTJ.TJ032,0)*-1 as 本幣未稅金額
                     FROM PURTJ as PURTJ  
                     Left JOIN PURTI as PURTI On PURTJ.TJ001=PURTI.TI001 and PURTJ.TJ002=PURTI.TI002 
                     Left JOIN INVMB as INVMB On PURTJ.TJ004=INVMB.MB001 
                     Left JOIN PURMA as PURMA On PURTI.TI004=PURMA.MA001
                     WHERE {0}
                     And left(PURTI.TI014,6) between '{1}' and '{2}'
                     ORDER BY left(INVMB.MB006,2) asc,PURMA.MA085 asc,PURTI.TI014 asc,PURTJ.TJ001 asc,PURTJ.TJ002 asc"
                    , cond_PURTJ, str_date_e_ym.Substring(0, 4) + "01", str_date_e_ym);

                MyCode.Sql_dgv(sql_str_PURTJ, dt_PURTJ, dgv_PURTJ);

                //IPSTF	S/I 資料單身檔、CMSMG	 幣別匯率檔單身 (當月明細)彙總表
                string sql_str_IPS_Now_Sum = String.Format(
                    @"select 產品別,關係人代號,關係人簡稱,sum(原物料進貨淨額) as 原物料進貨淨額,sum(商品進貨淨額) as 商品進貨淨額 from 
                        (--======20200587 在途倉明細(精簡)=====================
                        select 
                             left(MB006,2) as 產品別
                            ,PURMA.MA085 as 關係人代號
                            ,(case PURMA.MA085 when '82000' then '東洋科美' else PURMA.MA002 End) as 關係人簡稱
                            ,ROUND((case when MB005 in ('03','04') then (TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)) else 0 End),0) as 原物料進貨淨額
                            ,ROUND((case when MB005 in ('03','04') then 0 else (TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)) End),0) as 商品進貨淨額
                         from IPSTF
                            left join IPSTE on TE001 = TF001
                            left join PURMA on MA001 = TE003
                            left join PURTD on TD001 = TF002 and TD002 = TF003 and TD003 = TF004
                            left join PURTC on TC001 = TD001 and TC002 = TD002 
                            left join INVMB on MB001 = TD004
                        where left(TF015,6) = '{0}'
                        ) b
                        group by 關係人代號,關係人簡稱,產品別"
                         , str_date_e_ym);

                MyCode.Sql_dgv(sql_str_IPS_Now_Sum, dt_IPS_Now_Sum, dgv_IPS_Now_Sum);

                //IPSTF	S/I 資料單身檔、CMSMG	 幣別匯率檔單身 (當月明細)
                //在途資訊，keyin S/I資料建立作業；依單身[確認預交日]換算匯率，基本上為上月底為準

                string sql_str_IPS_Now = String.Format(
                         @"select PURMA.MA085 as 關係人代號,TE003 as 廠商代號
                            , (case when MB005 in ('03','04') then '原料' else '商品' End) as 進貨類別
                            ,TF001 as SI單號,	TF002 as 採購單別,	TF003 as 採購單號,	TF004 as 採購序號
                            ,	TF005 as 入庫庫別,	TF006 as 採購單價一,	TF009 as 採購數量,	TF010 as 採購金額一
                            ,	TF015 as 確認預交日,	TD004 as 品號,	TD005 as 品名,	TD006 as 規格
                            ,	TD009 as 單位,	TC005 as 交易幣別,	TC006 as 原匯率
                            , (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC) as 新匯率
                            , (select MA003 from INVMA where MA001 = '1' and MA002 = MB005) as 會計別名稱, left(MB006,2) as 產品別
                            , ROUND((TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)),0) as 本幣未稅金額 
                             from IPSTF
                                left join IPSTE on TE001 = TF001
                                left join PURMA on MA001 = TE003
                                left join PURTD on TD001 = TF002 and TD002 = TF003 and TD003 = TF004
                                left join PURTC on TC001 = TD001 and TC002 = TD002 
                                left join INVMB on MB001 = TD004
                            where left(TF015,6) = '{0}'"
                        , str_date_e_ym);

                MyCode.Sql_dgv(sql_str_IPS_Now, dt_IPS_Now, dgv_IPS_Now);

                //IPSTF	S/I 資料單身檔、CMSMG	 幣別匯率檔單身 (去年明細)減彙總表
                string sql_str_IPS_Last_Sum = String.Format(
                    @"select 產品別,關係人代號,關係人簡稱,sum(原物料進貨淨額) as 原物料進貨淨額,sum(商品進貨淨額) as 商品進貨淨額 from 
                        (--======20200507 去年底在途倉明細(縮減)=====================
                        select 
                             left(MB006,2) as 產品別
                            ,PURMA.MA085 as 關係人代號
                            ,(case PURMA.MA085 when '82000' then '東洋科美' else PURMA.MA002 End) as 關係人簡稱
                            ,ROUND((case when MB005 in ('03','04') then -(TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)) else 0 End),0) as 原物料進貨淨額
                            ,ROUND((case when MB005 in ('03','04') then 0 else -(TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)) End),0) as 商品進貨淨額
                         from IPSTF
                            left join IPSTE on TE001 = TF001
                            left join PURMA on MA001 = TE003
                            left join PURTD on TD001 = TF002 and TD002 = TF003 and TD003 = TF004
                            left join PURTC on TC001 = TD001 and TC002 = TD002 
                            left join INVMB on MB001 = TD004
                        where left(TF015,6) = '{0}'
                        ) b
                        group by 關係人代號,關係人簡稱,產品別"
                         , str_date_e_lasty + "12");

                MyCode.Sql_dgv(sql_str_IPS_Last_Sum, dt_IPS_Last_Sum, dgv_IPS_Last_Sum);

                //IPSTF	S/I 資料單身檔、CMSMG	 幣別匯率檔單身 (去年明細)減
                //在途資訊，keyin S/I資料建立作業；依單身[確認預交日]換算匯率，基本上為上月底為準
                string sql_str_IPS_Last = String.Format(
                         @"select PURMA.MA085 as 關係人代號,TE003 as 廠商代號
                            , (case when MB005 in ('03','04') then '原料' else '商品' End) as 進貨類別
                            ,TF001 as SI單號,	TF002 as 採購單別,	TF003 as 採購單號,	TF004 as 採購序號
                            ,	TF005 as 入庫庫別,	TF006 as 採購單價一,	TF009 as 採購數量,	TF010 as 採購金額一
                            ,	TF015 as 確認預交日,	TD004 as 品號,	TD005 as 品名,	TD006 as 規格
                            ,	TD009 as 單位,	TC005 as 交易幣別,	TC006 as 原匯率
                            , (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC) as 新匯率
                            , (select MA003 from INVMA where MA001 = '1' and MA002 = MB005) as 會計別名稱, left(MB006,2) as 產品別
                            , ROUND((TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)),0) as 本幣未稅金額 
                             from IPSTF
                                left join IPSTE on TE001 = TF001
                                left join PURMA on MA001 = TE003
                                left join PURTD on TD001 = TF002 and TD002 = TF003 and TD003 = TF004
                                left join PURTC on TC001 = TD001 and TC002 = TD002 
                                left join INVMB on MB001 = TD004
                            where left(TF015,6) = '{0}'"
                        , str_date_e_lasty + "12");

                MyCode.Sql_dgv(sql_str_IPS_Last, dt_IPS_Last, dgv_IPS_Last);

            }
            BtnTrue();

            tabControl1.SelectedIndex = 0;
        }

        private void fm_Acc_5b_Load(object sender, EventArgs e)
        {
            txt_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
            string filder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            txt_path.Text = filder;

            //20210118 財務林姿刪提出，去除347單別
            //20210720 財務林姿刪提出，去除進退貨條件式【(case when PURTG.TG005 = N'TVS' AND PURTG.TG007 = N'JPY' then 0 else 1 end=1)】
            cond_5b = @"彙總[進貨明細表、在途當月(加)、在途去年(減)]";
            //cond_PURTH = @"((PURTH.TH030 = N'Y') AND (PURTH.TH001 in(N'340',N'345',N'349',N'C340')) 
            //         AND (PURMA.MA085 <> '') AND (case when PURTG.TG005 = N'TVS' AND PURTG.TG007 = N'JPY' then 0 else 1 end=1))";
            //cond_PURTJ = @"((PURTJ.TJ020 = N'Y') AND (PURTJ.TJ001 in(N'350',N'355',N'359',N'C350'))
            //         AND (PURMA.MA085 <> '') AND (case when PURTI.TI004 = N'TVS' AND PURTI.TI006 = N'JPY' then 0 else 1 end=1))";

            cond_PURTH = @"((PURTH.TH030 = N'Y') AND (PURTH.TH001 in(N'340',N'345',N'349',N'C340')) 
                     AND (PURMA.MA085 <> ''))";
            cond_PURTJ = @"((PURTJ.TJ020 = N'Y') AND (PURTJ.TJ001 in(N'350',N'355',N'359',N'C350'))
                     AND (PURMA.MA085 <> ''))";

            cond_IPS_Now = @"";
            cond_IPS_Last = @"取[結束日期]抓取年份，換算去年年底，例如：201912";

            txterr.Text = string.Format(
                @"1.在途取[結束日期]抓取月份，例如：2021/02/28，將抓取[2021/02]資訊，
進退貨日期，抓取[結束日期]年初，例如：2021/02/28，將抓取[202101-202102]資訊
2.日期變更後，先前查詢資料須重新查詢，若無查詢，禁止Excel轉出。
3.Excel轉出後包含明細，程式自動開啟該報表。
4.查詢條件：
========   5b明細彙總表  ===========
{0}
========   進貨單明細  PURTH ===========
{1}
========   退貨單明細  PURTJ ===========
{2}
====  S/I建立作業 IPSTF 在途當月(加) ====
====  S/I建立作業 IPSTF 在途去年(減) ====
{3}
", cond_5b, cond_PURTH, cond_PURTJ, cond_IPS_Last);
        
    }

        private void btn_5b_Click(object sender, EventArgs e)
        {
            save_as_5b = txt_path.Text.ToString().Trim() + "\\" + str_date_e_ym.Substring(0, 4) + "01" + "-" + str_date_e_ym.Substring(4, 2) + @"關聯方進貨淨額明細5b_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
            path = txt_path.Text.ToString().Trim();
            //判断文件夹是否存在
            bool directoryExist = Directory.Exists(path);
            if (!directoryExist)
            {
                //创建
                Directory.CreateDirectory(path);//关联创建所有层级
            }
            //判断文件是否存在
            bool fileExist = File.Exists(save_as_5b);
            if (fileExist)
            {
                DialogResult myResult = MessageBox.Show
                ("是否取代原本檔案?", "檔案已存在", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (myResult == DialogResult.Yes)
                {
                    //按了是
                    //删除文件
                    File.Delete(save_as_5b);
                }
                else if (myResult == DialogResult.No)
                {
                    //按了否
                    return;
                }
            }

            BtnFalse();

            using (XLWorkbook wb_5b = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel_5b))
                {
                    var ws = templateWB.Worksheet("進貨淨額彙總表");
                    var ws2 = templateWB.Worksheet("進貨明細表");
                    var ws3 = templateWB.Worksheet("退貨明細表");
                    var ws4 = templateWB.Worksheet("當月在途彙總表");
                    var ws5 = templateWB.Worksheet("當月在途存貨(加)");
                    var ws6 = templateWB.Worksheet("去年在途彙總表");
                    var ws7 = templateWB.Worksheet("去年12月在途存貨(減)");

                    ws.CopyTo(wb_5b, "進貨淨額彙總表");
                    ws2.CopyTo(wb_5b, "進貨明細表");
                    ws3.CopyTo(wb_5b, "退貨明細表");
                    ws4.CopyTo(wb_5b, "當月在途彙總表");
                    ws5.CopyTo(wb_5b, "當月在途存貨(加)");
                    ws6.CopyTo(wb_5b, "去年在途彙總表");
                    ws7.CopyTo(wb_5b, "去年12月在途存貨(減)");
                }

                var wsheet_5b = wb_5b.Worksheet("進貨淨額彙總表");
                var wsheet_PURTH = wb_5b.Worksheet("進貨明細表");
                var wsheet_PURTJ = wb_5b.Worksheet("退貨明細表");
                var wsheet_IPS_Now_Sum = wb_5b.Worksheet("當月在途彙總表");
                var wsheet_IPS_Now = wb_5b.Worksheet("當月在途存貨(加)");
                var wsheet_IPS_Last_Sum = wb_5b.Worksheet("去年在途彙總表");
                var wsheet_IPS_Last = wb_5b.Worksheet("去年12月在途存貨(減)");

                //== 5b進貨彙總表.進貨明細表.當月在途存貨(加).去年12月在途存貨(減) =======
                ERP_DTInputExcel(wsheet_5b, dt_5b,6,1, str_date_e_ym.Substring(0, 4) + "01");
                ERP_DTInputExcel(wsheet_PURTH, dt_PURTH,5,1, str_date_e_ym.Substring(0, 4) + "01");
                ERP_DTInputExcel(wsheet_PURTJ, dt_PURTJ, 5, 1, str_date_e_ym.Substring(0, 4) + "01");
                ERP_DTInputExcel(wsheet_IPS_Now_Sum, dt_IPS_Now_Sum, 6, 1, str_date_e_ym);
                ERP_DTInputExcel(wsheet_IPS_Now, dt_IPS_Now,5,1, str_date_e_ym);
                ERP_DTInputExcel(wsheet_IPS_Last_Sum, dt_IPS_Last_Sum, 6, 1, str_date_e_ym);
                ERP_DTInputExcel(wsheet_IPS_Last, dt_IPS_Last,5,1, str_date_e_ym);

                wb_5b.SaveAs(save_as_5b);

                //打开文件
                System.Diagnostics.Process.Start(save_as_5b);
            }
            BtnTrue();

        }

        private void btn_PUR_Click(object sender, EventArgs e)
        {
            save_as_PUR = txt_path.Text.ToString().Trim() + "\\" + str_date_e_ym.Substring(0, 4) + "01" + "-" + str_date_e_ym.Substring(4, 2) + @"關聯方進貨淨額明細5b(進退貨明細表)_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
            path = txt_path.Text.ToString().Trim();

            //判断文件夹是否存在
            bool directoryExist = Directory.Exists(path);
            if (!directoryExist)
            {
                //创建
                Directory.CreateDirectory(path);//关联创建所有层级
            }
            //判断文件是否存在
            bool fileExist = File.Exists(save_as_PUR);
            if (fileExist)
            {
                DialogResult myResult = MessageBox.Show
                ("是否取代原本檔案?", "檔案已存在", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (myResult == DialogResult.Yes)
                {
                    //按了是
                    //删除文件
                    File.Delete(save_as_PUR);
                }
                else if (myResult == DialogResult.No)
                {
                    //按了否
                    return;
                }
            }

            BtnFalse();

            using (XLWorkbook wb_5b = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel_5b))
                {
                    var ws1 = templateWB.Worksheet("進貨明細表");
                    var ws2 = templateWB.Worksheet("退貨明細表");

                    ws1.CopyTo(wb_5b, "進貨明細表");
                    ws2.CopyTo(wb_5b, "退貨明細表");
                }

                var wsheet_PURTH = wb_5b.Worksheet("進貨明細表");
                var wsheet_PURTJ = wb_5b.Worksheet("退貨明細表");

                //== 進貨明細表 =======
                ERP_DTInputExcel(wsheet_PURTH, dt_PURTH,5,1, str_date_e_ym.Substring(0, 4) + "01");
                //== 退貨明細表 =======
                ERP_DTInputExcel(wsheet_PURTJ, dt_PURTJ, 5, 1, str_date_e_ym.Substring(0, 4) + "01");

                wb_5b.SaveAs(save_as_PUR);

                //打开文件
                    System.Diagnostics.Process.Start(save_as_PUR);
            }
            BtnTrue();
        }

        private void btn_IPS_Click(object sender, EventArgs e)
        {
            save_as_IPS = txt_path.Text.ToString().Trim() + "\\" + str_date_s_ym + "-" + str_date_e_ym.Substring(4, 2) + @"關聯方進貨淨額明細5b(在途倉明細)_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
            path = txt_path.Text.ToString().Trim();

            //判断文件夹是否存在
            bool directoryExist = Directory.Exists(path);
            if (!directoryExist)
            {
                //创建
                Directory.CreateDirectory(path);//关联创建所有层级
            }
            //判断文件是否存在
            bool fileExist = File.Exists(save_as_IPS);
            if (fileExist)
            {
                DialogResult myResult = MessageBox.Show
                ("是否取代原本檔案?", "檔案已存在", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (myResult == DialogResult.Yes)
                {
                    //按了是
                    //删除文件
                    File.Delete(save_as_IPS);
                }
                else if (myResult == DialogResult.No)
                {
                    //按了否
                    return;
                }
            }

            BtnFalse();

            using (XLWorkbook wb_5b = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel_5b))
                {
                    var ws3 = templateWB.Worksheet("當月在途存貨(加)");
                    var ws4 = templateWB.Worksheet("去年12月在途存貨(減)");

                    ws3.CopyTo(wb_5b, "當月在途存貨(加)");
                    ws4.CopyTo(wb_5b, "去年12月在途存貨(減)");
                }

                var wsheet_IPS_Now = wb_5b.Worksheet("當月在途存貨(加)");
                var wsheet_IPS_Last = wb_5b.Worksheet("去年12月在途存貨(減)");

                //== 當月在途存貨(加).去年12月在途存貨(減) =======
                ERP_DTInputExcel(wsheet_IPS_Now, dt_IPS_Now,5,1, str_date_s_ym);
                ERP_DTInputExcel(wsheet_IPS_Last, dt_IPS_Last,5,1, str_date_s_ym);

                wb_5b.SaveAs(save_as_IPS);

                //打开文件
                System.Diagnostics.Process.Start(save_as_IPS);
            }
            BtnTrue();
        }

        private void ERP_DTInputExcel(ClosedXML.Excel.IXLWorksheet wsheet, DataTable dt, int i,int j, string str_date)
        {
            //int i = 0;
            int j_def = j;

            if  (dt.TableName.Substring(0, 2) == "當月")
            {
                wsheet.Cell(2, 2).Value = str_date + "-" + str_date; //查詢月份區間
                wsheet.Cell(3, 2).Style.NumberFormat.Format = "@";
                wsheet.Cell(3, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期
            }
            else if (dt.TableName.Substring(0, 2) == "去年")
            {
                wsheet.Cell(2, 2).Value = str_date_e_lasty + "12" + "-" + str_date_e_lasty + "12"; //查詢月份區間
                wsheet.Cell(3, 2).Style.NumberFormat.Format = "@";
                wsheet.Cell(3, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期
            }
            else
            {
                wsheet.Cell(2, 2).Value = str_date + "-" + str_date_e_ym; //查詢月份區間
                wsheet.Cell(3, 2).Style.NumberFormat.Format = "@";
                wsheet.Cell(3, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期
            }


            foreach (DataRow row in dt.Rows)
            {
                j = j_def;
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
                        case "關係人代號" :
                        case "品種別":
                        case "進貨單別":
                        case "進貨單號":
                        case "退貨單別":
                        case "退貨單號":
                        case "SI單號":
                        case "採購單別":
                        case "採購單號":
                        case "採購序號":
                        case "入庫庫別":
                        case "確認預交日":
                            wsheet.Cell(i, j).Style.NumberFormat.Format = "@";
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
                        case "原料進貨金額":
                        case "商品進貨金額":
                        case "原料退貨金額":
                        case "商品退貨金額":
                        case "原料進貨淨額":
                        case "商品進貨淨額":
                        case "原物料進貨金額":
                        case "原物料進貨淨額":
                            wsheet.Cell(i, j).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                            break;
                        case "銷貨數量":
                        case "銷貨數":
                        case "銷退數":
                        case "進貨數量":
                        case "退貨數量":
                        case "採購數量":
                            wsheet.Cell(i, j).Style.NumberFormat.Format = "#,##0.000;[RED](#,##0.000)";
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
                        case "原幣退貨金額":
                        case "原幣進貨淨額":
                        case "原幣退貨淨額":
                            wsheet.Cell(i, j).Style.NumberFormat.Format = "#,##0.00;[RED](#,##0.00)";
                            break;
                        default:
                            break;
                    }
                    wsheet.Cell(i, j).Value = row[row_num];
                    row_num++;
                    j++;
                }
                i++;
            }
        }

        bool IsToForm1 = false; //紀錄是否要回到Form1
        protected override void OnClosing(CancelEventArgs e) //在視窗關閉時觸發
        {
            DialogResult dr = MessageBox.Show("\"是\"回到主畫面 \r\n \"否\"關閉程式", "是否要關閉程式"
                , MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                IsToForm1 = true;
            }
            else if (dr == DialogResult.Cancel)
            {

            }

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
