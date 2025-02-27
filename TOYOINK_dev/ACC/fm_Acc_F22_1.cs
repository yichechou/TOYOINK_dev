﻿using System;
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
using DocumentFormat.OpenXml.Spreadsheet;

namespace TOYOINK_dev
{
    /*
     * 20210202 完成
     * 注意事項：幣別欄位需去除前後空白
     * 20210602 洪淑雯提出，(1)類型"轉帳"，原幣出帳金額錯誤，查看修正[原幣出帳.本幣出帳]，判別為'81' 有值其餘為0，
     * ,(case MQ003 when '82' then TG008 else 0 End) as	原幣入帳金額
	   ,(case MQ003 when '81' then TG008 else 0 End) as	原幣出帳金額  <-修正
       ,(case MQ003 when '82' then TG019 else 0 End) as 本幣入帳金額
	   ,(case MQ003 when '81' then TG019 else 0 End) as 本幣出帳金額  <-修正
        
        (2)//外幣存款月底重評價表
            20210602 修正
            (case MQ003 when '81' then TG008*-1 when '82' then TG008 end) as 原幣 ,
					(case MQ003 when '81' then TG019*-1 when '82' then TG019 end) as 本幣
                    from NOTTG
                    left join NOTTF on TF001 = TG001 and TF002 = TG002
                    left join CMSMQ on MQ001 = TF001

       20210902 修改連線方式，改為參考MyClass
       20210902 洪淑雯提出，評價表內的[小計.總計]改為取小數點2位
       20220418 財務林姿刪提出，
       1.新增[小口現金],銀行代號1105101 零用金及週轉金-台南 =SUM(SUMIFS('明細帳(評價前)'!H:H,'明細帳(評價前)'!B:B,E42,'明細帳(評價前)'!A:A,R42))
       2.新增[現金・普通預金・当座預金]、[定期預金]、[借入金]合計
       3.定期預金,新增換算金額(殘證)、換算金額(差額)以及合計
       NOTMA.UDF01 = '1'，改為 NOTMA.UDF01 in ('1','2')
       
       20231017 財務 林姿刪提出，資訊系統(開發改善)提案書(AMIS003)-2310170001
        一、新增分頁-"明細分類帳"
            1.EXCEL 新增 "明細分類帳"分頁，參考5a明細分類帳語法，條件為【ML009(摘要) like '%應兌現票據%' and ML006(科目代號) = '1102041'】

        二、分頁-F22-1_銀行口座一覧表TAST 調整
            1.EXCEL 新增【批准者/承認者欄位】及【作成 / 時點】日期自動寫入。
            2.A43欄位，新增【應兌現票據欄位】，公式【=SUM(SUMIFS(明細分類帳!F:F,明細分類帳!A:A,R43))】；R43【銀行代號-1102041】。
            3.J45.K45 欄位，修改【現金・普通預金・当座預金】，計算公式【=J26+F42(小口現金)+F43(應兌現票據)】、【=K26+G42+G43】。

        三、明細帳(評價前)、明細帳(評價後)
            1.調整分頁內主管簽核欄位位置，原本設定為C50，改為【C50-總經理：、G50-財務部部長：、K50-製表人：】。

        四、EXCEL整體
            1.日期區間格式，改為查詢日期區間【2023/09/01~2023/09/30】。
            2.顯示比例改為80%。
            3.字體設定為【微軟正黑體】。
            4.調整欄位寬度。
      * 20240513 更新NuGet套件後出現錯誤，修改程式碼加入【(ClosedXML.Excel.XLCellValue)】；再次修改，刪除前面修改，結尾加入【.ToString()】
      * 20240524 因轉出數值為文字，再次修改程式，改為數值
      * 20240527 財務 余若玫 提出，銀行帳號數值呈現[科學記數方式]，增加判別式
      * 20240528 再次調整，0開頭或關係人代號...等改為文字欄位
     */
    public partial class fm_Acc_F22_1 : Form
    {
        public MyClass MyCode;
        月曆 fm_月曆;
        string save_as_F22_1_Month = "", temp_excel_F22_1 = "";
        string createday = DateTime.Now.ToString("yyyy/MM/dd");
        int opencode = 0;

        string str_date_s, str_date_m_s, str_date_ym_s;
        string str_date_e, str_date_m_e, str_date_ym_e, str_date_y_e;

        string defaultfilePath = "";

        DateTime date_s, date_e;

        DataTable dt_SGL_Before = new DataTable();    //銀行存款明細帳(評價前)
        DataTable dt_SGL_After = new DataTable();     //銀行存款明細帳(評價後)
        DataTable dt_ADFOR = new DataTable();         //外幣存款月底重評價表
        DataTable dt_SGL_Detail = new DataTable();    //銀行存款明細帳細項
        DataTable dt_ACTML = new DataTable();    //明細分類帳


        public fm_Acc_F22_1()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();

            //MyCode.strDbCon = MyCode.strDbConLeader;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConLeader;

            MyCode.strDbCon = MyCode.strDbConA01A;
            //this.sqlConnection1.ConnectionString = MyCode.strDbConA01A;

            temp_excel_F22_1 = @"\\192.168.128.219\Conductor\Company\MIS自開發主檔\會計報表公版\F22-1_銀行口座一覧表TAST_temp.xlsx";
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

        private void fm_Acc_F22_1_Load(object sender, EventArgs e)
        {

            txt_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
            string filder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            txt_path.Text = filder;

            txterr.Text = string.Format(
                @"1.取[結束]抓取月份，例如：2021/01/29，將抓取[2021/01]資訊。
2.日期變更後，先前查詢資料須重新查詢，若無查詢，禁止Excel轉出。
3.Excel轉出後包含明細，程式自動開啟該報表。
4.銀行帳號建立作業 NOTMA，新增自訂欄位UDF01，
【1使用中.2零用金.3變更代號.0銷戶】，抓取【1,2】
5.查詢條件(幣別需去除前後空白)：
======== 銀行存款明細帳(評價前) ===========
以NOTMA為主檔串查合併 月統計表(NOTLA)及明細表(CT_F22_1_SGLDT_After_Temp) 
來源類型 <> '匯兌損益' 
NOTMA.UDF01 in ('1','2')，
需再加入【銀行月統計檔NOTLA】並需扣除【匯兌損益調整單身檔(NOTTQ)】
======== 銀行存款明細帳(評價後) ===========
NOTMA.UDF01 in ('1','2')，
來自【存款提款:銀行存款存提單頭檔NOTTF+轉帳:銀行存款存提單身檔NOTTG+應收兌現:應收票據單頭檔NOTTC+匯兌損益:匯兌損益調整單身檔NOTTQ】合併
加入【銀行月統計檔NOTLA】
======== 外幣存款月底重評價表 ===========
幣別 <> 'NTD'
======== 明細分類帳 ===========
ML009 like '%應兌現票據%' and ML006 = '1102041'
");

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

        private void DtAndDgvClear()
        {
            //清除
            //檢查tempdb內,是否有此暫存資料表
            string sql_str_Detel_Temp_Table = String.Format(@"
                    delete CT_F22_1_SGLDT_After_Temp");
            MyCode.sqlExecuteNonQuery(sql_str_Detel_Temp_Table, "S2008X64");

            dt_SGL_Before.Clear();   //銀行存款明細帳(評價前)
            dt_SGL_After.Clear();    //銀行存款明細帳(評價後)
            dt_ADFOR.Clear();        //外幣存款月底重評價表
            dt_SGL_Detail.Clear();   //銀行存款明細帳細項
            dt_ACTML.Clear();   //明細分類帳


            dgv_SGL_Before.DataSource = null;
            dgv_SGL_After.DataSource = null;
            dgv_ADFOR.DataSource = null;
            dgv_SGL_Detail.DataSource = null;
            dgv_ACTML.DataSource = null;

            BtnFalse();
        }

        private void BtnFalse()
        {
            btn_ToE_F22_1_M.Enabled = false;
        }
        private void BtnTrue()
        {
            btn_ToE_F22_1_M.Enabled = true;
        }
        private void btn_search_Click(object sender, EventArgs e)
        {
            //dgvData.Columns[3].DefaultCellStyle.Format = "###,###,##0.00";

            if (MyClass.DateIntervalCheck(txt_date_s, txt_date_e) is false)
            {
                return;
            }

            DtAndDgvClear();

            str_date_s = txt_date_s.Text.Trim();
            str_date_ym_s = txt_date_s.Text.Trim().Substring(0, 6);
            str_date_e = txt_date_e.Text.Trim();
            str_date_ym_e = txt_date_e.Text.Trim().Substring(0, 6);
            //str_date_y_e = txt_date_e.Text.Trim().Substring(0, 4);


            //銀行存款明細帳細項_評價後及匯入暫存表CT_F22_1_SGLDT_After_Temp
            string sql_str_Insert_CT_F22_1_SGLDT_After_Temp = String.Format(@"
                    INSERT INTO CT_F22_1_SGLDT_After_Temp (銀行代號,幣別,銀行簡稱,銀行帳號,日期,匯率,原幣入帳金額,原幣出帳金額
                    ,本幣入帳金額,本幣出帳金額,廠商代號,廠商簡稱,單據號碼,付款銀行,來源類型,備註)
                    select 銀行代號,幣別,銀行簡稱,銀行帳號,日期,匯率,原幣入帳金額,原幣出帳金額
                    ,本幣入帳金額,本幣出帳金額,廠商代號,廠商簡稱,單據號碼,付款銀行,來源類型,備註
                    from
                        (select TG005 as	銀行代號, rtrim(TG017) as	幣別
                        ,(select MA002 from NOTMA where MA001 = TG005) as 銀行簡稱 
                        ,(select MA004 from NOTMA where MA001 = TG005) as 銀行帳號
                        , TF003 as 日期, TG018 as 匯率
                        ,(case MQ003 
                        when '82' then TG008
                        else 0 End) as	原幣入帳金額
						,(case MQ003 
                        when '81' then TG008
                        else 0 End) as	原幣出帳金額
                        ,(case MQ003 
                        when '82' then TG019
                        else 0 End)	as 本幣入帳金額
						,(case MQ003 
                        when '81' then TG019
                        else 0 End)	as 本幣出帳金額
                        , '' as 廠商代號, '' as 廠商簡稱
                        , (Rtrim(TG001) + '-' + Rtrim(TG002)) as 單據號碼
                        , '' as 付款銀行, '轉帳'	as 來源類型, TF007 as 備註
                        from NOTTG
                        left join NOTTF on TF001 = TG001 and TF002 = TG002
                        left join CMSMQ on MQ001 = TF001
                        where  TG010 = 'Y' and TG011 = '1' and TG004 = '2' and TF003 like '{0}%'
                    union all
                        select TQ004 as	銀行代號, rtrim(TQ005) as	幣別
                        ,(select MA002 from NOTMA where MA001 = TQ004) as 銀行簡稱 
                        ,(select MA004 from NOTMA where MA001 = TQ004) as 銀行帳號
                        , TP009 as 日期, TQ006 as 匯率
                        , 0 as	原幣入帳金額, 0 as	原幣出帳金額
                        ,(case when TQ009 > 0 then TQ009
                        else 0 End)	as 本幣入帳金額
                        ,(case when TQ009 < 0 then TQ009*-1 
                        else 0 End)	as 本幣出帳金額
                        , '' as 廠商代號, '' as 廠商簡稱
                        , (Rtrim(TQ001) + '-' + Rtrim(TQ002)) as 票據號碼
                        , '' as 付款銀行, '匯兌損益' as 來源類型, TQ010 as 備註
                         from NOTTQ
                        left join NOTTP on TP001 = TQ001 and TP002 = TQ002
                        where  TP010 = 'Y' and TP009 like '{0}%'
                    union all
                        select TC025 as	銀行代號, rtrim(TC002) as	幣別
                        ,(select MA002 from NOTMA where MA001 = TC025) as 銀行簡稱 
                        ,(select MA004 from NOTMA where MA001 = TC025) as 銀行帳號
                        , TD003 as 日期, TD005 as 匯率, TC003 as	原幣入帳金額, 0 as	原幣出帳金額
                        , TC003*TD005 as	本幣入帳金額, 0 as	本幣出帳金額
                        , TC013 as 廠商代號, TC014 as 廠商簡稱, TC001 as 單據號碼, TC008 as 付款銀行
                        , '應收兌現' as 來源類型, TC024 as 備註
                         from NOTTC
                        left join NOTTD on TD001 = TC001
                        where  TD004 = '6' and TD003 like '{0}%'
                    union all
                        select TF004 as	銀行代號, rtrim(TF005) as	幣別
                        ,(select MA002 from NOTMA where MA001 = TF004) as 銀行簡稱 
                        ,(select MA004 from NOTMA where MA001 = TF004) as 銀行帳號
                        , TF003 as 日期, TF006 as 匯率
                        ,(case MQ003 
                        when '81' then TF013
                        else 0 End) as	原幣入帳金額
                        ,(case MQ003 
                        when '82' then TF013
                        else 0 End) as	原幣出帳金額
                        ,(case MQ003 
                        when '81' then TF014
                        else 0 End)	as 本幣入帳金額
                        ,(case MQ003 
                        when '82' then TF014
                        else 0 End)	as 本幣出帳金額
                        , '' as 廠商代號, '' as 廠商簡稱
                        , (Rtrim(TF001) + '-' + Rtrim(TF002)) as 單據號碼, '' as 付款銀行
                        ,(case MQ003 
                        when '81' then '提款'
                        when '82' then '存款' End)	as 來源類型, TF007 as 備註
                         from NOTTF
                        left join CMSMQ on MQ001 = TF001
                        where  TF010 = 'Y' and TF003 like '{0}%') SGL_Detail_After
                    left join NOTMA on  MA001 = 銀行代號
                    where NOTMA.UDF01 in ('1','2')", str_date_ym_e);
            MyCode.sqlExecuteNonQuery(sql_str_Insert_CT_F22_1_SGLDT_After_Temp, "S2008X64");

            string sql_str_CT_F22_1_SGLDT_After_Temp = String.Format(@"
                    select * from CT_F22_1_SGLDT_After_Temp order by 銀行代號,日期");
            MyCode.Sql_dgv(sql_str_CT_F22_1_SGLDT_After_Temp, dt_SGL_Detail, dgv_SGL_Detail);

            //銀行存款明細帳(評價前)
            string sql_str_SGL_Before = String.Format(@"select  銀行代號,幣別,銀行簡稱,銀行帳號
                        ,sum(原幣期初餘額) as 原幣期初餘額
                        ,(case when sum(原幣入帳金額) is Null then '0'
                            else sum(原幣入帳金額) End) as 原幣入帳金額
                        ,(case when sum(原幣出帳金額) is Null then '0'
                            else sum(原幣出帳金額) End) as 原幣出帳金額
                        ,(sum(原幣期初餘額) + sum(原幣入帳金額) - sum(原幣出帳金額)) as 原幣期末餘額
                        ,sum(本幣期初餘額) as 本幣期初餘額
                        ,(case when sum(本幣入帳金額) is Null then '0'
                            else sum(本幣入帳金額) End) as 本幣入帳金額
                        ,(case when sum(本幣出帳金額) is Null then '0'
                            else sum(本幣出帳金額) End) as 本幣出帳金額
                        ,(sum(本幣期初餘額)+ sum(本幣入帳金額) - sum(本幣出帳金額)) as 本幣期末餘額
                        from (select LA001 as 銀行代號 , rtrim(LA003) as 幣別, MA002 as 銀行簡稱, MA004 as 銀行帳號
                            , LA004 as 原幣期初餘額, '0' as 原幣入帳金額, '0' as 原幣出帳金額, '0' as 原幣期末餘額
                            , LA006 as 本幣期初餘額, '0' as 本幣入帳金額, '0' as 本幣出帳金額, '0' as 本幣期末餘額
                             from NOTLA
                            left join NOTMA on MA001 = LA001
                             where NOTMA.UDF01 in ('1','2') and LA002 = '{0}'
                        union all
                            select  銀行代號, 幣別, 銀行簡稱, 銀行帳號
                            ,'0' as 原幣期初餘額, 原幣入帳金額, 原幣出帳金額,'0' as 原幣期末餘額
                            ,'0' as 本幣期初餘額, 本幣入帳金額, 本幣出帳金額,'0' as 本幣期末餘額
                             from CT_F22_1_SGLDT_After_Temp where 來源類型 <> '匯兌損益') SGL_Before
                        group by 銀行代號,幣別,銀行簡稱,銀行帳號", str_date_ym_e);
            MyCode.Sql_dgv(sql_str_SGL_Before, dt_SGL_Before, dgv_SGL_Before);

            //銀行存款明細帳(評價後)
            string sql_str_SGL_After = String.Format(@"select  銀行代號,幣別,銀行簡稱,銀行帳號
                        ,sum(原幣期初餘額) as 原幣期初餘額
                        ,(case when sum(原幣入帳金額) is Null then '0'
                            else sum(原幣入帳金額) End) as 原幣入帳金額
                        ,(case when sum(原幣出帳金額) is Null then '0'
                            else sum(原幣出帳金額) End) as 原幣出帳金額
                        ,(sum(原幣期初餘額) + sum(原幣入帳金額) - sum(原幣出帳金額)) as 原幣期末餘額
                        ,sum(本幣期初餘額) as 本幣期初餘額
                        ,(case when sum(本幣入帳金額) is Null then '0'
                            else sum(本幣入帳金額) End) as 本幣入帳金額
                        ,(case when sum(本幣出帳金額) is Null then '0'
                            else sum(本幣出帳金額) End) as 本幣出帳金額
                        ,(sum(本幣期初餘額)+ sum(本幣入帳金額) - sum(本幣出帳金額)) as 本幣期末餘額
                        from (select LA001 as 銀行代號 , rtrim(LA003) as 幣別, MA002 as 銀行簡稱, MA004 as 銀行帳號
                            , LA004 as 原幣期初餘額, '0' as 原幣入帳金額, '0' as 原幣出帳金額, '0' as 原幣期末餘額
                            , LA006 as 本幣期初餘額, '0' as 本幣入帳金額, '0' as 本幣出帳金額, '0' as 本幣期末餘額
                             from NOTLA
                            left join NOTMA on MA001 = LA001
                             where NOTMA.UDF01 in ('1','2') and LA002 = '{0}'
                        union all
                            select  銀行代號, 幣別, 銀行簡稱, 銀行帳號
                            ,'0' as 原幣期初餘額, 原幣入帳金額, 原幣出帳金額,'0' as 原幣期末餘額
                            ,'0' as 本幣期初餘額, 本幣入帳金額, 本幣出帳金額,'0' as 本幣期末餘額
                             from CT_F22_1_SGLDT_After_Temp) SGL_After
                        group by 銀行代號,幣別,銀行簡稱,銀行帳號", str_date_ym_e);
            MyCode.Sql_dgv(sql_str_SGL_After, dt_SGL_After, dgv_SGL_After);


            //外幣存款月底重評價表
            /*20210602 修正
            (case MQ003 when '81' then TG008*-1 when '82' then TG008 end) as 原幣 ,
					(case MQ003 when '81' then TG019*-1 when '82' then TG019 end) as 本幣
                    from NOTTG
                    left join NOTTF on TF001 = TG001 and TF002 = TG002
                    left join CMSMQ on MQ001 = TF001
            */
            string sql_str_ADFOR = String.Format(@"select 幣別, 重估匯率, 銀行代號, 銀行簡稱
                , 原幣存款金額, 平均匯率, 本幣存款金額, 重估本幣金額
                ,(case when (重估本幣金額-本幣存款金額) > 0 then (重估本幣金額-本幣存款金額) else 0 End)	as 匯兌收益
                ,(case when (本幣存款金額-重估本幣金額) > 0 then (本幣存款金額-重估本幣金額) else 0 End)	as 匯兌損失,'' as 淨損益
                 from (
                    select 幣別
                    ,cast(round((select Top 1 MG003 from CMSMG where MG002 <= '{0}' and MG001 = 幣別 order by MG002 desc),4)as numeric(20,4)) as 重估匯率
                    ,銀行代號 ,(select MA002 from NOTMA where MA001 = 銀行代號) as 銀行簡稱 
                    ,sum(原幣) as 原幣存款金額 ,cast(round(ISNULL(sum(本幣)/NULLIF(sum(原幣), 0), 0),4)as numeric(20,4)) as 平均匯率 ,sum(本幣) as 本幣存款金額
                    ,ROUND((cast(round((select Top 1 MG003 from CMSMG where MG002 <= '{0}' and MG001 = 幣別 order by MG002 desc),4)as numeric(20,4)) * sum(原幣)),0)  as 重估本幣金額
                    ,'' as 匯兌收益 ,'' as 匯兌損失 ,'' as 淨損益
                    from 
                        ((select TF004 as 銀行代號 ,left(TF003,6) as 存款年月 ,rtrim(TF005) as 幣別 
                        ,sum((case MQ003 
                        when '81' then (TF013) 
                        when '82' then (TF013*-1) Else TF013 End)) as 原幣
                        ,sum((case MQ003 
                        when '81' then (TF014) 
                        when '82' then (TF014*-1) Else TF014 End)) as 本幣
                        from NOTTF
                        left join CMSMQ on MQ001 = TF001
                        where TF010 = 'Y' and  TF005 <> 'NTD'  and TF003 like '{1}%'
                        group by TF004,TF005,left(TF003,6))
                union all
                    (select LA001 as 銀行代號 ,LA002 as 存款年月 ,rtrim(LA003) as 幣別 ,LA004 as 原幣  ,LA006 as 本幣
                    from NOTLA
                    where LA003 <> 'NTD' and LA002 = '{1}')
                union all
                    (select TG005 as 銀行代號 ,left(TF003,6) as 存款年月 ,rtrim(TG017) as 幣別 ,
                    (case MQ003 when '81' then TG008*-1 when '82' then TG008 end) as 原幣 ,
					(case MQ003 when '81' then TG019*-1 when '82' then TG019 end) as 本幣
                    from NOTTG
                    left join NOTTF on TF001 = TG001 and TF002 = TG002
                    left join CMSMQ on MQ001 = TF001
                    where TG010 = 'Y' and TG017 <> 'NTD' and TG011 = '1' and TG004 = '2' and TF003 like '{1}%')) NOT_sum 
                    group by 幣別,銀行代號) NOT_Fin
                where 本幣存款金額 <> 0
                order by 幣別,銀行代號", str_date_e, str_date_ym_e);
            MyCode.Sql_dgv(sql_str_ADFOR, dt_ADFOR, dgv_ADFOR);

            //dt_sumtotal(dt_SGL_Before, "本幣期末餘額");

            //明細分類帳
            string sql_str_ACTML = String.Format(@"select * from (
                                  select ML006 as 科目編號 
                                    ,(select MA003 from ACTMA where MA001 = ACTML.ML006) as 科目名稱
                                    ,SUBSTRING(ML002,1,6) as 傳票年月 ,ML003+'-'+ML004+' -'+ML005 as 傳票編號
                                    ,ML009 as 摘要
                                    ,(case ML007 when '1' then ML008 else 0 end) as 本幣借方金額
                                    ,(case ML007 when '-1' then ML008 else 0 end)  as 本幣貸方金額 
                                    ,(case MA007 
	                                    when '1' then ((case ML007 when '1' then ML008 else 0 end)-(case ML007 when '-1' then ML008 else 0 end)) 
	                                    when '-1' then ((case ML007 when '-1' then ML008 else 0 end)-(case ML007 when '1' then ML008 else 0 end)) else 0 end ) as 貸借金額
                                    from ACTML
	                                    left JOIN ACTMA on ACTMA.MA001 = ACTML.ML006
                                        left JOIN ACTTB on ACTTB.TB001 = ACTML.ML003 and ACTTB.TB002 = ACTML.ML004 and ACTTB.TB003 = ACTML.ML005
                                    where (
	                                (ML009 like '%應兌現票據%' and ML006 = '1102041') 
	                                )) ACTML_ALL
                            where 傳票年月 like '{0}'
                            order by 傳票年月,科目編號", str_date_ym_e);
            MyCode.Sql_dgv(sql_str_ACTML, dt_ACTML, dgv_ACTML);

            BtnTrue();
        }
        //private DataTable dt_sumtotal(DataTable dt, string str_Total)
        //{
        //    DataRow dr = dt.NewRow();
        //    dr[0] = "合計";
        //    dr[str_Total] = "=sum(L1:L21)";//无效的聚合函数 Sum()和类型 String 的用法 数据库中的数据类型必须是数字
        //    dt.Rows.Add(dr);
        //    return dt;
        //}

        private void btn_ToE_F22_1_M_Click(object sender, EventArgs e)
        {
            BtnFalse();

            using (XLWorkbook wb_F22_1_Month = new XLWorkbook())
            {
                using (var templateWB = new XLWorkbook(temp_excel_F22_1))
                {
                    var ws = templateWB.Worksheet("F22-1_銀行口座一覧表TAST");
                    var ws2 = templateWB.Worksheet("明細帳(評價前)");
                    var ws3 = templateWB.Worksheet("明細帳(評價後)");
                    var ws4 = templateWB.Worksheet("評價表");
                    var ws5 = templateWB.Worksheet("明細帳細項");
                    var ws6 = templateWB.Worksheet("明細分類帳");

                    ws.CopyTo(wb_F22_1_Month, "F22-1_銀行口座一覧表TAST");
                    ws2.CopyTo(wb_F22_1_Month, "明細帳(評價前)");
                    ws3.CopyTo(wb_F22_1_Month, "明細帳(評價後)");
                    ws4.CopyTo(wb_F22_1_Month, "評價表");
                    ws5.CopyTo(wb_F22_1_Month, "明細帳細項");
                    ws6.CopyTo(wb_F22_1_Month, "明細分類帳");
                }

                var wsheet_F22_1_m = wb_F22_1_Month.Worksheet("F22-1_銀行口座一覧表TAST");
                var wsheet_SGL_Before = wb_F22_1_Month.Worksheet("明細帳(評價前)");
                var wsheet_SGL_After = wb_F22_1_Month.Worksheet("明細帳(評價後)");
                var wsheet_ADFOR = wb_F22_1_Month.Worksheet("評價表");
                var wsheet_SGL_Detail = wb_F22_1_Month.Worksheet("明細帳細項");
                var wsheet_ACTML = wb_F22_1_Month.Worksheet("明細分類帳");

                //=== F22-1_銀行口座一覧表TAST ==========================================
                //wsheet_F22_1_m.Cell(2, 1).Value = "月份區間:" + str_date_ym_s + "~" + str_date_ym_e; //查詢月份區間
                //wsheet_F22_1_m.Cell(3, 1).Value = "製表日期:" + DateTime.Now.ToString("yyyy/MM/dd"); //會計年度

                ////== 明細帳(評價前).明細帳(評價後).評價表 =======
                ///ERP_DTInputExcel(wsheet_8aCOPTH, dt_8aCOPTH, str_date_y_e + "01");
            
                wsheet_F22_1_m.Cell(1, 16).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期
                wsheet_F22_1_m.Cell(2, 16).Value = DateTime.ParseExact(str_date_e, "yyyyMMdd", null).ToString("yyyy/MM/dd");
                wsheet_F22_1_m.Cells().Style.Font.FontName = "微軟正黑體";
                ERP_DTInputExcel(wsheet_SGL_Before, dt_SGL_Before, 5, 1, str_date_s, "", "本幣期末餘額");
                ERP_DTInputExcel(wsheet_SGL_After, dt_SGL_After, 5, 1, str_date_s, "", "本幣期末餘額");
                ERP_DTInputExcel(wsheet_ADFOR, dt_ADFOR, 5, 1, str_date_s, "幣別", "原幣存款金額");
                ERP_DTInputExcel(wsheet_SGL_Detail, dt_SGL_Detail, 5, 1, str_date_s, "", "");
                ERP_DTInputExcel(wsheet_ACTML, dt_ACTML, 5, 1, str_date_s, "", "");
                //ERP_DTInputExcel(wsheet_ADFOR, dt_ADFOR, str_date_ym_s, "幣別", "原幣存款金額;本幣存款金額;重估本幣金額;匯兌損失;淨(損)益");

                save_as_F22_1_Month = txt_path.Text.ToString().Trim() + "\\" + str_date_ym_e + @"_F22-1_銀行口座一覧表TAST_" + DateTime.Now.ToString("yyyyMMdd") + @".xlsx";
                wb_F22_1_Month.SaveAs(save_as_F22_1_Month);

                //打开文件
                if (opencode != 1)
                {
                    System.Diagnostics.Process.Start(save_as_F22_1_Month);
                }
            }
            BtnTrue();
        }

        private void ERP_DTInputExcel(ClosedXML.Excel.IXLWorksheet wsheet, DataTable dt,int i_col, int j_row, string str_date,string str_SubTotal,string str_Total)
        {
            int i = 0;
            int rows_count_dt = dt.Rows.Count;
            int col_count_dt = dt.Columns.Count;
            string str_SubTotal_Name = "";

            

            DateTime date_sheet_s = DateTime.ParseExact(str_date, "yyyyMMdd", null);
            DateTime date_sheet_e = DateTime.ParseExact(str_date_e, "yyyyMMdd", null);

            wsheet.Cell(2, 2).Value = date_sheet_s.ToString("yyyy/MM/dd") + "~" + date_sheet_e.ToString("yyyy/MM/dd"); //查詢月份區間
            wsheet.Cell(3, 2).Style.NumberFormat.Format = "@";
            wsheet.Cell(3, 2).Value = DateTime.Now.ToString("yyyy/MM/dd"); //製表日期


            foreach (DataRow row in dt.Rows)
            {
                int j = 0;

                if (str_SubTotal.Length > 0 && str_Total.Length > 0)
                {
                    if (str_SubTotal_Name.ToString() != "" && row[str_SubTotal].ToString() != str_SubTotal_Name.ToString())
                    {
                        wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet.Cell(i + i_col, 2).Value = str_SubTotal_Name;
                        wsheet.Cell(i + i_col, 4).Value = "小計";
                        //20210902 洪淑雯提出，評價表內的[小計.總計]改為取小數點2位
                        //wsheet.Range("E" + (i + i_col) + ":K" + (i + i_col)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet.Cell(i + i_col, 5).Style.NumberFormat.Format = "#,##0.00";
                        wsheet.Range("G" + (i + i_col) + ":K" + (i + i_col)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet.Cell(i + i_col, 5).FormulaA1 = "=SUMIFS(E:E,$A:$A,\"" + str_SubTotal_Name + "\")";
                        wsheet.Cell(i + i_col, 7).FormulaA1 = "=SUMIFS(G:G,$A:$A,\"" + str_SubTotal_Name + "\")";
                        wsheet.Cell(i + i_col, 8).FormulaA1 = "=SUMIFS(H:H,$A:$A,\"" + str_SubTotal_Name + "\")";
                        wsheet.Cell(i + i_col, 9).FormulaA1 = "=SUMIFS(I:I,$A:$A,\"" + str_SubTotal_Name + "\")";
                        wsheet.Cell(i + i_col, 10).FormulaA1 = "=SUMIFS(J:J,$A:$A,\"" + str_SubTotal_Name + "\")";
                        //wsheet.SheetView.ZoomScale = 80;
                        i++;
                    }
                }

                foreach (DataColumn Column in dt.Columns)
                {
                    string format = null;

                    switch (Column.ColumnName.ToString())
                    {
                        case "銀行帳號":
                        case "銀行代號":
                        case "科目編號":
                        case "傳票年月":
                            wsheet.Cell(i + i_col, j + j_row).Style.NumberFormat.Format = "@";
                            break;
                        case "本幣期初餘額":
                        case "本幣入帳金額":
                        case "本幣出帳金額":
                        case "本幣期末餘額":
                        case "本幣存款金額":
                        case "本幣借方金額":
                        case "本幣貸方金額":
                        case "貸借金額":
                        case "重估本幣金額":
                        case "匯兌收益":
                        case "匯兌損失":
                            wsheet.Cell(i + i_col, j + j_row).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                            break;
                        case "原幣進貨金額":
                        case "原幣期初餘額":
                        case "原幣入帳金額":
                        case "原幣出帳金額":
                        case "原幣期末餘額":
                        case "原幣存款金額":
                        case "單位製費成本":
                            wsheet.Cell(i + i_col, j + j_row).Style.NumberFormat.Format = "#,##0.00;[RED](#,##0.00)";
                            break;
                        case "匯率":
                        case "重估匯率":
                        case "平均匯率":
                            wsheet.Cell(i + i_col, j + j_row).Style.NumberFormat.Format = "#,##0.0000";
                            break;
                        default:
                            break;
                    }
                    //20240513 更新NuGet套件後出現錯誤，修改程式碼加入【(ClosedXML.Excel.XLCellValue)】；再次修改，刪除前面修改，結尾加入【.ToString()】
                    //wsheet.Cell(i + i_col, j + j_row).Value = row[j];
                    //wsheet.Cell(i + i_col, j + j_row).Value = row[j].ToString();
                    //20240524 因轉出數值為文字，再次修改程式，改為數值
                    //20240527 財務 余若玫 提出，銀行帳號數值呈現[科學記數方式]，增加判別式
                    //20240528 再次調整，0開頭或關係人代號...等改為文字欄位
                    // 設置單元格的值和格式
                    SetCellValueAndFormat(wsheet, i + i_col, j + j_row, row[j], format);
                    j++;
                }

                if (str_SubTotal.Length > 0 && str_Total.Length > 0)
                {
                    str_SubTotal_Name = row[str_SubTotal].ToString().Trim();
                }

                if ((rows_count_dt - 1) == dt.Rows.IndexOf(row)) //資料列結尾運算
                {
                    if (str_SubTotal.Length == 0 && str_Total.Length > 0)
                    //if (wsheet.ToString() == "明細帳(評價前)" || wsheet.ToString() == "明細帳(評價後)")
                    {
                        i++;
                        wsheet.Range("A" + (i + i_col) + ":L" + (i + i_col)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet.Range("A" + (i + i_col) + ":L" + (i + i_col)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet.Range("A" + (i + i_col) + ":L" + (i + i_col)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet.Cell(i + i_col, col_count_dt - 2).Value = "小計";
                        wsheet.Cell(i + i_col, j + j_row - 1).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet.Cell(i + i_col, j + j_row - 1).FormulaA1 = "=sum(L" + i_col + ":L" + (i + i_col - 1) + ")";
                        //wsheet.SheetView.ZoomScale = 80;

                    }
                    //if (wsheet.ToString() == "評價表")
                    if (str_SubTotal.Length > 0 && str_Total.Length > 0)
                    {
                        i++;
                        wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet.Cell(i + i_col, 2).Value = str_SubTotal_Name;
                        wsheet.Cell(i + i_col, 4).Value = "小計";
                        //20210902 洪淑雯提出，評價表內的[小計.總計]改為取小數點2位
                        //wsheet.Range("E" + (i + i_col) + ":K" + (i + i_col)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)"
                        wsheet.Cell(i + i_col, 5).Style.NumberFormat.Format = "#,##0.00";
                        wsheet.Range("G" + (i + i_col) + ":K" + (i + i_col)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet.Cell(i + i_col, 5).FormulaA1 = "=SUMIFS(E:E,$A:$A,\"" + str_SubTotal_Name + "\")";
                        wsheet.Cell(i + i_col, 7).FormulaA1 = "=SUMIFS(G:G,$A:$A,\"" + str_SubTotal_Name + "\")";
                        wsheet.Cell(i + i_col, 8).FormulaA1 = "=SUMIFS(H:H,$A:$A,\"" + str_SubTotal_Name + "\")";
                        wsheet.Cell(i + i_col, 9).FormulaA1 = "=SUMIFS(I:I,$A:$A,\"" + str_SubTotal_Name + "\")";
                        wsheet.Cell(i + i_col, 10).FormulaA1 = "=SUMIFS(J:J,$A:$A,\"" + str_SubTotal_Name + "\")";

                        i++;
                        wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet.Range("A" + (i + i_col) + ":K" + (i + i_col)).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet.Cell(i + i_col, 4).Value = "總計";
                        //20210902 洪淑雯提出，評價表內的[小計.總計]改為取小數點2位
                        //wsheet.Range("E" + (i + i_col) + ":K" + (i + i_col)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)"
                        wsheet.Cell(i + i_col, 5).Style.NumberFormat.Format = "#,##0.00";
                        wsheet.Range("G" + (i + i_col) + ":K" + (i + i_col)).Style.NumberFormat.Format = "#,##0_);[RED](#,##0)";
                        wsheet.Cell(i + i_col, 5).FormulaA1 = "=SUMIFS(E:E,$D:$D,\"小計\")";
                        wsheet.Cell(i + i_col, 7).FormulaA1 = "=SUMIFS(G:G,$D:$D,\"小計\")";
                        wsheet.Cell(i + i_col, 8).FormulaA1 = "=SUMIFS(H:H,$D:$D,\"小計\")";
                        wsheet.Cell(i + i_col, 9).FormulaA1 = "=SUMIFS(I:I,$D:$D,\"小計\")";
                        wsheet.Cell(i + i_col, 10).FormulaA1 = "=SUMIFS(J:J,$D:$D,\"小計\")";
                        wsheet.Cell(i + i_col, 11).FormulaA1 = "=I" + (i + i_col) +  "-J" + (i + i_col);
                        

                    }
                }
               
                i++;
                wsheet.Cells().Style.Font.FontName = "微軟正黑體";
                wsheet.SheetView.ZoomScale = 80;
            }

        }
        //20240528 再次調整，0開頭或關係人代號...等改為文字欄位
        void SetCellValueAndFormat(IXLWorksheet sheet, int rowIndex, int colIndex, object value, string format = null)
        {
            // 檢查是否有指定格式
            if (format != null)
            {
                sheet.Cell(rowIndex, colIndex).Style.NumberFormat.Format = format;
            }

            // 判斷 value 是否為數字，並且是否需要保留前導零
            if (value is string strValue && strValue.StartsWith("0") && double.TryParse(strValue, out _))
            {
                // 如果 value 是以 "0" 開頭的字符串且可以解析為數字，則保留字符串形式
                sheet.Cell(rowIndex, colIndex).Value = strValue;
            }
            else if (double.TryParse(value.ToString(), out double numericValue))
            {
                // 如果 value 可以解析為數字，且大於 1e10，則設置為字符串值
                if (numericValue > 1e10)
                {
                    sheet.Cell(rowIndex, colIndex).Value = value.ToString();
                }
                else
                {
                    // 否則，設置為數字值
                    sheet.Cell(rowIndex, colIndex).Value = numericValue;
                }
            }
            else
            {
                // 否則，設置為字符串值
                sheet.Cell(rowIndex, colIndex).Value = value.ToString();
            }
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
/*
 * CT_F22_1_SGLDT_After_Temp 自訂明細暫存表
資料行名稱	資料類型	允許NULL
銀行代號	nvarchar(10)	Checked
幣別	nvarchar(4)	Checked
銀行簡稱	nvarchar(30)	Checked
銀行帳號	nvarchar(30)	Checked
日期	nvarchar(8)	Checked
匯率	numeric(18, 4)	Checked
原幣入帳金額	numeric(18, 2)	Checked
原幣出帳金額	numeric(18, 2)	Checked
本幣入帳金額	numeric(18, 0)	Checked
本幣出帳金額	numeric(18, 0)	Checked
廠商代號	nvarchar(10)	Checked
廠商簡稱	nvarchar(30)	Checked
單據號碼	nvarchar(20)	Checked
付款銀行	nvarchar(10)	Checked
來源類型	nvarchar(20)	Checked
備註	nvarchar(255)	Checked
 * 
 */
