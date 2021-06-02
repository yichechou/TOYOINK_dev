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
    //20200923 在途查詢加入取整數 ROUND((TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)),0) as 新本幣金額
    public partial class fm_Package7b : Form
    {
        public MyClass MyCode;
        月曆 fm_月曆;

        string str_enter = ((char)13).ToString() + ((char)10).ToString();

        DataTable dt_7B = new DataTable();  //權利金-全部
        DataTable dt_INV = new DataTable();  //折讓-全部
        DataTable dt_IPS = new DataTable();  //權利金-光阻

        string createday = DateTime.Now.ToString("yyyy/MM/dd");

        string str_date_s;
        string str_date_e, str_date_e_ym;

        string defaultfilePath = "";
        string path,fileNameWithExtension;

        DateTime date_s,date_e;

        public fm_Package7b()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();
            MyCode.strDbCon = "packet size=4096;user id=pwuser;password=sqlmis003;data source=192.168.128.219;persist security info=False;initial catalog=A01A;";
            //MyCode.strDbCon = "packet size=4096;user id=yj.chou;password=yjchou3369;data source=192.168.128.219;persist security info=False;initial catalog=Leader;";
            //temp_excel = @"\\192.168.128.219\Company\會計\權利金與折讓報表\權利金與折讓報表_temp.xlsx";
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

        private void fm_Package7b_FormClosed(object sender, FormClosedEventArgs e)
        {
            Environment.Exit(Environment.ExitCode);
        }


        private void txt_date_s_TextChanged(object sender, EventArgs e)
        {
            btn_7B.Enabled = false;
            btn_INV.Enabled = false;
            btn_IPS.Enabled = false;

            if (dgv_7B.DataSource != null)
            {
                dt_7B.Clear();
                dt_INV.Clear();
                dt_IPS.Clear();

                dgv_7B.DataSource = null;
                dgv_INV.DataSource = null;
                dgv_IPS.DataSource = null;
            }
        }

        private void txt_date_e_TextChanged(object sender, EventArgs e)
        {
            btn_7B.Enabled = false;
            btn_INV.Enabled = false;
            btn_IPS.Enabled = false;

            if (dgv_7B.DataSource != null)
            {
                dt_7B.Clear();
                dt_INV.Clear();
                dt_IPS.Clear();

                dgv_7B.DataSource = null;
                dgv_INV.DataSource = null;
                dgv_IPS.DataSource = null;
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


        private void fm_Package7b_Load(object sender, EventArgs e)
        {

            txt_date_s.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddMonths(-1).ToString("yyyyMMdd");
            txt_date_e.Text = DateTime.Parse(DateTime.Now.ToString("yyyy-MM-01")).AddDays(-1).ToString("yyyyMMdd");
            string filder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            txt_path.Text = filder;
            
            txterr.Text = string.Format(
                @"1.取[結束]抓取月份，例如：2020/02/29，將抓取[2020/02]資訊。
2.日期變更後，先前查詢資料須重新查詢，若無查詢，禁止Excel轉出。
3.Excel轉出後，程式自動開啟該報表，亦可各別轉出明細。
");
                
        }


        private void btn_search_Click(object sender, EventArgs e)
        {
            btn_7B.Enabled = false;
            btn_INV.Enabled = false;
            btn_IPS.Enabled = false;

            date_s = DateTime.ParseExact(txt_date_s.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);
            date_e = DateTime.ParseExact(txt_date_e.Text.Trim(), "yyyyMMdd", CultureInfo.InvariantCulture);

            if (date_s > date_e)
            {
                MessageBox.Show("請修改日期區間", "日期格式錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (dgv_7B.DataSource != null)
            {
                dt_7B.Clear();
                dt_INV.Clear();
                dt_IPS.Clear();

                dgv_7B.DataSource = null;
                dgv_INV.DataSource = null;
                dgv_IPS.DataSource = null;
            }

            str_date_e_ym = txt_date_e.Text.Trim().Substring(0, 6);

            string sql_str_7B = String.Format(
                @" SELECT LMB006 as 產品別整,PURMAMA085 as 關係人代號,PURMAMA002 as 供應商簡稱,IMB006 as 存貨科目,sum(END_MONEY) as 期末庫存額,YM as 庫存月份
                 from 
                 (SELECT left(INVMB.MB006,2) as LMB006
                 ,PURMA.MA085 as PURMAMA085
                 ,(case PURMA.MA085 when '82000' then '東洋科美' else PURMA.MA002 End) as PURMAMA002
                 ,(case left(INVMB.MB006,2)when '52' then '5'else '1' end) as IMB006
                 ,ROUND((TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)),0) as END_MONEY
                 ,left(IPSTF.TF015,6) as YM
                  FROM IPSTF
                  Left JOIN PURTD as PURTD On IPSTF.TF002=PURTD.TD001 and IPSTF.TF003=PURTD.TD002 and IPSTF.TF004=PURTD.TD003 
                  Left JOIN IPSTE as IPSTE On IPSTE.TE001=IPSTF.TF001 
                  Left JOIN PURTC as PURTC On PURTC.TC001=PURTD.TD001 AND PURTC.TC002=PURTD.TD002 
                  Left JOIN INVMB as INVMB On PURTD.TD004=INVMB.MB001 
                  Left JOIN PURMA as PURMA On IPSTE.TE003=PURMA.MA001
                union all
                 SELECT left(INVMB.MB006,2) as LMB006
                 ,PURMA.MA085 as PURMAMA085
                 ,(case PURMA.MA085 when '82000' then '東洋科美' else PURMA.MA002 End) as PURMAMA002
                ,(case left(INVMB.MB006,2)when '52' then '5'else '1' end) as IMB006
                ,(INVLC.LC005+INVLC.LC007-INVLC.LC009-INVLC.LC011+INVLC.LC013+INVLC.LC015-INVLC.LC017+INVLC.LC019+INVLC.LC021-INVLC.LC023-INVLC.LC025) as END_MONEY
                ,INVLC.LC002 as YM
                 FROM INVLC 
                 Left JOIN INVMB as INVMB On INVLC.LC001=INVMB.MB001 
                 Left JOIN PURMA as PURMA On INVMB.MB032=PURMA.MA001
                 WHERE INVMB.MB032 <> '' and PURMA.MA085 <> ''
                 and ((INVLC.LC005+INVLC.LC007-INVLC.LC009-INVLC.LC011+INVLC.LC013+INVLC.LC015-INVLC.LC017+INVLC.LC019+INVLC.LC021-INVLC.LC023-INVLC.LC025) <> 0)) B
                 where YM ='{0}'
                 group by  LMB006,PURMAMA085,PURMAMA002,IMB006,YM
                 order by PURMAMA085"
                , str_date_e_ym);

            MyCode.Sql_dgv(sql_str_7B, dt_7B, dgv_7B);

            //INVLC	品號每月統計單身
            //期末庫存=期初+本期入庫-本期銷貨-本期領料+本期轉撥(入)+本期調整(入)-本期出庫+本期銷退+本期退料-本期轉撥(出)-本期調整(出)
            string sql_str_INV = String.Format(
                 @"SELECT left(INVMB.MB006,2) as 產品別整,PURMA.MA085 as 關係人代號,INVMB.MB032 as 供應商代號,PURMA.MA002 as 供應商簡稱
                    ,(case left(INVMB.MB006,2)when '52' then '5'else '1' end) as 存貨科目
                    ,INVLC.LC005+INVLC.LC007-INVLC.LC009-INVLC.LC011+INVLC.LC013+INVLC.LC015-INVLC.LC017+INVLC.LC019+INVLC.LC021-INVLC.LC023-INVLC.LC025 as 期末庫存額
                    ,INVLC.LC001 as 品號,INVMB.MB002 as 品名,INVMB.MB006 as 原產品別,INVLC.LC002 as 庫存月份
                     FROM INVLC as INVLC  Left JOIN INVMB as INVMB On INVLC.LC001=INVMB.MB001 Left JOIN PURMA as PURMA On INVMB.MB032=PURMA.MA001
                     WHERE ((INVMB.MB032 <> '') and (PURMA.MA085 <> '') 
                    and ((INVLC.LC005+INVLC.LC007-INVLC.LC009-INVLC.LC011+INVLC.LC013+INVLC.LC015-INVLC.LC017+INVLC.LC019+INVLC.LC021-INVLC.LC023-INVLC.LC025) <> 0))
                    and INVLC.LC002 = '{0}' order by left(INVMB.MB006,2)"
                , str_date_e_ym);

            MyCode.Sql_dgv(sql_str_INV, dt_INV, dgv_INV);

            //IPSTF	S/I 資料單身檔、CMSMG	 幣別匯率檔單身
            //在途資訊，keyin S/I資料建立作業；依單身[確認預交日]換算匯率，基本上為上月底為準

            string sql_str_IPS = String.Format(
                     @"SELECT IPSTE.TE002 as SI單據日期,IPSTF.TF001 as SI單號,PURMA.MA085 as 關係人代號,IPSTE.TE003 as 供應商代號
                         ,PURMA.MA002 as 供應商簡稱,(case left(INVMB.MB006,2)when '52' then '5'else '1' end) as 存貨科目
                         ,PURTD.TD004 as 品號,PURTD.TD005 as 品名,IPSTF.TF010 as 原幣金額,IPSTF.TF015 as 確認交貨日
                         ,(select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC) as 新匯率
                         ,ROUND((TF010 * (select top 1 MG006 from CMSMG where MG001 = TC005 and MG002 < TF015 ORDER BY MG002 DESC)),0) as 新本幣金額
                         ,PURTD.TD009 as 單位,PURTC.TC005 as 幣別,INVMB.MB006 as 產品別原代號
                         ,(select MA003 from INVMA as INVMA2 where MA001= '2' and INVMB.MB006 = INVMA2.MA002) as 產品別名稱,IPSTF.TF002 as 採購單別
                         ,IPSTF.TF003 as 採購單號,IPSTF.TF004 as 採購序號,IPSTF.TF005 as 庫別
                          FROM IPSTF as IPSTF  Left JOIN PURTD as PURTD On IPSTF.TF002=PURTD.TD001 and IPSTF.TF003=PURTD.TD002 and IPSTF.TF004=PURTD.TD003 
                          Left JOIN IPSTE as IPSTE On IPSTE.TE001=IPSTF.TF001 Left JOIN PURTC as PURTC On PURTC.TC001=PURTD.TD001 AND PURTC.TC002=PURTD.TD002 
                          Left JOIN INVMB as INVMB On PURTD.TD004=INVMB.MB001 Left JOIN PURMA as PURMA On IPSTE.TE003=PURMA.MA001
                          where IPSTF.TF015 like '{0}%'
                          order by IPSTF.TF001"
                    , str_date_e_ym);

            MyCode.Sql_dgv(sql_str_IPS, dt_IPS, dgv_IPS);

            tabControl1.SelectedIndex = 0;
            btn_7B.Enabled = true;
            btn_INV.Enabled = true;
            btn_IPS.Enabled = true;
        }
        /*
        //TODO:確認路徑下Excel是否存在，取消或取代
        private void CheckExcelFile(string path, string fileNameWithExtension)
        {
            //判断文件夹是否存在
            bool directoryExist = Directory.Exists(path);
            if (!directoryExist)
            {
                //创建
                Directory.CreateDirectory(path);//关联创建所有层级
            }
            //判断文件是否存在
            bool fileExist = File.Exists(path + "\\" + fileNameWithExtension);
            if (fileExist)
            {
                DialogResult myResult = MessageBox.Show
                ("是否取代原本檔案?", "檔案已存在", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (myResult == DialogResult.Yes)
                {
                    //按了是
                    //删除文件
                    File.Delete(path + "\\" + fileNameWithExtension);
                }
                else if (myResult == DialogResult.No)
                {
                    //按了否
                   return;
                }
            }
        }*/
        //TODO:7B轉出
        private void btn_7B_Click(object sender, EventArgs e)
        {
            // MyCode.ClosedXMLExportExcel(dgv_7B, txt_path.Text.ToString().Trim(), "關聯方期末存貨Package7B_" + str_date_e_ym + "_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx");

            path = txt_path.Text.ToString().Trim();
            fileNameWithExtension = "關聯方期末存貨Package7B_" + str_date_e_ym + "_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            //判断文件夹是否存在
            //判断文件夹是否存在
            bool directoryExist = Directory.Exists(path);
            if (!directoryExist)
            {
                //创建
                Directory.CreateDirectory(path);//关联创建所有层级
            }
            //判断文件是否存在
            bool fileExist = File.Exists(path + "\\" + fileNameWithExtension);
            if (fileExist)
            {
                DialogResult myResult = MessageBox.Show
                ("是否取代原本檔案?", "檔案已存在", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (myResult == DialogResult.Yes)
                {
                    //按了是
                    //删除文件
                    File.Delete(path + "\\" + fileNameWithExtension);
                }
                else if (myResult == DialogResult.No)
                {
                    //按了否
                    return;
                }
            }

            //保存成文件
            using (XLWorkbook wb = new XLWorkbook())
            {
                //wb.Worksheets.Add(dt_7B, "7B彙總表");
                
                var wsheet_7B = wb.Worksheets.Add("7B彙總表");
                MyCode.ERP_DTInputExcel(wsheet_7B, dt_7B, 2, 1, "0", "0");
                
                // wsheet_7B.Range("E:E").Style.NumberFormat.Format = "#,##0";

                //wb.Worksheets.Add(dt_INV, "庫存明細表");
                var wsheet_INV = wb.Worksheets.Add("庫存明細表");
                MyCode.ERP_DTInputExcel(wsheet_INV, dt_INV, 2, 1, "0", "0");

                var wsheet_IPS = wb.Worksheets.Add("在途明細表");

                int rows_count_IPS = dt_IPS.Rows.Count;
                int y = 1; int x = 1;
                string num_IPS = "";

                wsheet_IPS.Range("A" + y + ":T" + y).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                wsheet_IPS.Range("A" + y + ":T" + y).Style.Border.InsideBorder = XLBorderStyleValues.Medium;

                wsheet_IPS.Range("A" + y + ":T" + y).Style.Fill.BackgroundColor = XLColor.FromHtml("#E0E0E0");
                wsheet_IPS.Cell(y, 1).Value = "SI單據日期";
                wsheet_IPS.Cell(y, 2).Value = "SI單號";
                wsheet_IPS.Cell(y, 3).Value = "關係人代號";
                wsheet_IPS.Cell(y, 4).Value = "供應商代號";
                wsheet_IPS.Cell(y, 5).Value = "供應商簡稱";
                wsheet_IPS.Cell(y, 6).Value = "存貨科目";
                wsheet_IPS.Cell(y, 7).Value = "品號";
                wsheet_IPS.Cell(y, 8).Value = "品名";
                wsheet_IPS.Cell(y, 9).Value = "原幣金額";
                wsheet_IPS.Cell(y, 10).Value = "確認交貨日";
                wsheet_IPS.Cell(y, 11).Value = "新匯率";
                wsheet_IPS.Cell(y, 12).Value = "新本幣金額";
                wsheet_IPS.Cell(y, 13).Value = "單位";
                wsheet_IPS.Cell(y, 14).Value = "幣別";
                wsheet_IPS.Cell(y, 15).Value = "產品別原代號";
                wsheet_IPS.Cell(y, 16).Value = "產品別名稱";
                wsheet_IPS.Cell(y, 17).Value = "採購單別";
                wsheet_IPS.Cell(y, 18).Value = "採購單號";
                wsheet_IPS.Cell(y, 19).Value = "採購序號";
                wsheet_IPS.Cell(y, 20).Value = "庫別";

                y = y + 1;

                foreach (DataRow row in dt_IPS.Rows)
                {
                    if (num_IPS.ToString() != "" && row[1].ToString().Trim() != num_IPS.ToString())
                    {
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_IPS.Cell(y, 2).Style.NumberFormat.Format = "@";
                        wsheet_IPS.Cell(y, 2).Value = num_IPS;
                        wsheet_IPS.Cell(y, 3).Value = "小計";
                        wsheet_IPS.Cell(y, 9).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_IPS.Cell(y, 9).FormulaA1 = "=sum(I" + x + ":I" + (y - 1) + ")";
                        wsheet_IPS.Cell(y, 12).Style.NumberFormat.Format = "#,##0";
                        wsheet_IPS.Cell(y, 12).FormulaA1 = "=sum(L" + x + ":L" + (y - 1) + ")";

                        x = y + 1;
                        y++;
                    }

                    wsheet_IPS.Cell(y, 1).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 1).Value = row[0]; //SI單據日期
                    wsheet_IPS.Cell(y, 2).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 2).Value = row[1]; //SI單號
                    wsheet_IPS.Cell(y, 3).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 3).Value = row[2]; //關係人代號
                    wsheet_IPS.Cell(y, 4).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 4).Value = row[3]; //供應商代號
                    wsheet_IPS.Cell(y, 5).Value = row[4]; //供應商簡稱
                    wsheet_IPS.Cell(y, 6).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 6).Value = row[5]; //存貨科目
                    wsheet_IPS.Cell(y, 7).Value = row[6]; //品號
                    wsheet_IPS.Cell(y, 8).Value = row[7]; //品名
                    wsheet_IPS.Cell(y, 9).Style.NumberFormat.Format = "#,##0.00";
                    wsheet_IPS.Cell(y, 9).Value = row[8]; //原幣金額
                    wsheet_IPS.Cell(y, 10).Value = row[9]; //確認交貨日
                    wsheet_IPS.Cell(y, 11).Style.NumberFormat.Format = "#,##0.0000";
                    wsheet_IPS.Cell(y, 11).Value = row[10]; //新匯率
                    wsheet_IPS.Cell(y, 12).Style.NumberFormat.Format = "#,##0";
                    wsheet_IPS.Cell(y, 12).Value = row[11]; //新本幣金額
                    wsheet_IPS.Cell(y, 13).Value = row[12]; //單位
                    wsheet_IPS.Cell(y, 14).Value = row[13]; //幣別
                    wsheet_IPS.Cell(y, 15).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 15).Value = row[14]; //產品別原代號
                    wsheet_IPS.Cell(y, 16).Value = row[15]; //產品別名稱
                    wsheet_IPS.Cell(y, 17).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 17).Value = row[16]; //採購單別
                    wsheet_IPS.Cell(y, 18).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 18).Value = row[17]; //採購單號
                    wsheet_IPS.Cell(y, 19).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 19).Value = row[18]; //採購序號
                    wsheet_IPS.Cell(y, 20).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 20).Value = row[19]; //庫別

                    num_IPS = row[1].ToString().Trim();

                    if ((rows_count_IPS - 1) == dt_IPS.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        y++;
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_IPS.Cell(y, 2).Style.NumberFormat.Format = "@";
                        wsheet_IPS.Cell(y, 2).Value = num_IPS;
                        wsheet_IPS.Cell(y, 3).Value = "小計";
                        wsheet_IPS.Cell(y, 9).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_IPS.Cell(y, 9).FormulaA1 = "=sum(I" + x + ":I" + (y - 1) + ")";
                        wsheet_IPS.Cell(y, 12).Style.NumberFormat.Format = "#,##0";
                        wsheet_IPS.Cell(y, 12).FormulaA1 = "=sum(L" + x + ":L" + (y - 1) + ")";


                        x = y + 1;
                        y++;

                        wsheet_IPS.Cell(y, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_IPS.Cell(y, 3).Value = "總計";
                        wsheet_IPS.Cell(y, 9).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_IPS.Cell(y, 9).FormulaA1 = "=SUMIF(C:C,\"小計\",I:I)";
                        wsheet_IPS.Cell(y, 12).Style.NumberFormat.Format = "#,##0";
                        wsheet_IPS.Cell(y, 12).FormulaA1 = "=SUMIF(C:C,\"小計\",L:L)";
                    }
                    y++;
                }

                //自动调整列的宽度
                wb.Worksheet(1).Columns().AdjustToContents();
                wb.Worksheet(2).Columns().AdjustToContents();
                wb.Worksheet(3).Columns().AdjustToContents();

                //保存文件
                wb.SaveAs(path + "\\" + fileNameWithExtension);
            }
            //打开文件
            System.Diagnostics.Process.Start(path + "\\" + fileNameWithExtension);

        }

        //TODO:庫存明細表轉出
        private void btn_INV_Click(object sender, EventArgs e)
        {
            path = txt_path.Text.ToString().Trim();
            fileNameWithExtension = "關聯方期末存貨Package7B_" + str_date_e_ym + "庫存明細表_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            //判断文件夹是否存在
            bool directoryExist = Directory.Exists(path);
            if (!directoryExist)
            {
                //创建
                Directory.CreateDirectory(path);//关联创建所有层级
            }
            //判断文件是否存在
            bool fileExist = File.Exists(path + "\\" + fileNameWithExtension);
            if (fileExist)
            {
                DialogResult myResult = MessageBox.Show
                ("是否取代原本檔案?", "檔案已存在", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (myResult == DialogResult.Yes)
                {
                    //按了是
                    //删除文件
                    File.Delete(path + "\\" + fileNameWithExtension);
                }
                else if (myResult == DialogResult.No)
                {
                    //按了否
                    return;
                }
            }

            //保存成文件
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt_INV, "庫存明細表");

                //自动调整列的宽度
                wb.Worksheet(1).Columns().AdjustToContents();

                //保存文件
                wb.SaveAs(path + "\\" + fileNameWithExtension);
            }

            //打开文件
            System.Diagnostics.Process.Start(path + "\\" + fileNameWithExtension);
        }

        //TODO:在途明細表轉出
        private void btn_IPS_Click(object sender, EventArgs e)
        {
            path = txt_path.Text.ToString().Trim();
            fileNameWithExtension = "關聯方期末存貨Package7B_" + str_date_e_ym + "在途明細表_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";
            //判断文件夹是否存在
            bool directoryExist = Directory.Exists(path);
            if (!directoryExist)
            {
                //创建
                Directory.CreateDirectory(path);//关联创建所有层级
            }
            //判断文件是否存在
            bool fileExist = File.Exists(path + "\\" + fileNameWithExtension);
            if (fileExist)
            {
                DialogResult myResult = MessageBox.Show
                ("是否取代原本檔案?", "檔案已存在", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (myResult == DialogResult.Yes)
                {
                    //按了是
                    //删除文件
                    File.Delete(path + "\\" + fileNameWithExtension);
                }
                else if (myResult == DialogResult.No)
                {
                    //按了否
                    return;
                }
            }

            //保存成文件
            using (XLWorkbook wb_IPS = new XLWorkbook())
            {
                var wsheet_IPS = wb_IPS.Worksheets.Add("在途明細表");
                
                int rows_count_IPS = dt_IPS.Rows.Count;
                int y = 1; int x = 1;
                string num_IPS = "";

                wsheet_IPS.Range("A" + y + ":T" + y).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                wsheet_IPS.Range("A" + y + ":T" + y).Style.Border.InsideBorder = XLBorderStyleValues.Medium;

                wsheet_IPS.Range("A" + y + ":T" + y).Style.Fill.BackgroundColor = XLColor.FromHtml("#E0E0E0");
                wsheet_IPS.Cell(y, 1).Value = "SI單據日期";
                wsheet_IPS.Cell(y, 2).Value = "SI單號";
                wsheet_IPS.Cell(y, 3).Value = "關係人代號";
                wsheet_IPS.Cell(y, 4).Value = "供應商代號";
                wsheet_IPS.Cell(y, 5).Value = "供應商簡稱";
                wsheet_IPS.Cell(y, 6).Value = "存貨科目";
                wsheet_IPS.Cell(y, 7).Value = "品號";
                wsheet_IPS.Cell(y, 8).Value = "品名";
                wsheet_IPS.Cell(y, 9).Value = "原幣金額";
                wsheet_IPS.Cell(y, 10).Value = "確認交貨日";
                wsheet_IPS.Cell(y, 11).Value = "新匯率";
                wsheet_IPS.Cell(y, 12).Value = "新本幣金額";
                wsheet_IPS.Cell(y, 13).Value = "單位";
                wsheet_IPS.Cell(y, 14).Value = "幣別";
                wsheet_IPS.Cell(y, 15).Value = "產品別原代號";
                wsheet_IPS.Cell(y, 16).Value = "產品別名稱";
                wsheet_IPS.Cell(y, 17).Value = "採購單別";
                wsheet_IPS.Cell(y, 18).Value = "採購單號";
                wsheet_IPS.Cell(y, 19).Value = "採購序號";
                wsheet_IPS.Cell(y, 20).Value = "庫別";

                y =y + 1;

                foreach (DataRow row in dt_IPS.Rows)
                {
                    if (num_IPS.ToString() != "" && row[1].ToString().Trim() != num_IPS.ToString())
                    {
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_IPS.Cell(y, 2).Style.NumberFormat.Format = "@";
                        wsheet_IPS.Cell(y, 2).Value = num_IPS;
                        wsheet_IPS.Cell(y, 3).Value = "小計";
                        wsheet_IPS.Cell(y, 9).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_IPS.Cell(y, 9).FormulaA1 = "=sum(I" + x + ":I" + (y - 1) + ")";
                        wsheet_IPS.Cell(y, 12).Style.NumberFormat.Format = "#,##0";
                        wsheet_IPS.Cell(y, 12).FormulaA1 = "=sum(L" + x + ":L" + (y - 1) + ")";

                        x = y + 1;
                        y++;
                    }

                    wsheet_IPS.Cell(y, 1).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 1).Value = row[0]; //SI單據日期
                    wsheet_IPS.Cell(y, 2).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 2).Value = row[1]; //SI單號
                    wsheet_IPS.Cell(y, 3).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 3).Value = row[2]; //關係人代號
                    wsheet_IPS.Cell(y, 4).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 4).Value = row[3]; //供應商代號
                    wsheet_IPS.Cell(y, 5).Value = row[4]; //供應商簡稱
                    wsheet_IPS.Cell(y, 6).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 6).Value = row[5]; //存貨科目
                    wsheet_IPS.Cell(y, 7).Value = row[6]; //品號
                    wsheet_IPS.Cell(y, 8).Value = row[7]; //品名
                    wsheet_IPS.Cell(y, 9).Style.NumberFormat.Format = "#,##0.00";
                    wsheet_IPS.Cell(y, 9).Value = row[8]; //原幣金額
                    wsheet_IPS.Cell(y, 10).Value = row[9]; //確認交貨日
                    wsheet_IPS.Cell(y, 11).Style.NumberFormat.Format = "#,##0.0000";
                    wsheet_IPS.Cell(y, 11).Value = row[10]; //新匯率
                    wsheet_IPS.Cell(y, 12).Style.NumberFormat.Format = "#,##0";
                    wsheet_IPS.Cell(y, 12).Value = row[11]; //新本幣金額
                    wsheet_IPS.Cell(y, 13).Value = row[12]; //單位
                    wsheet_IPS.Cell(y, 14).Value = row[13]; //幣別
                    wsheet_IPS.Cell(y, 15).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 15).Value = row[14]; //產品別原代號
                    wsheet_IPS.Cell(y, 16).Value = row[15]; //產品別名稱
                    wsheet_IPS.Cell(y, 17).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 17).Value = row[16]; //採購單別
                    wsheet_IPS.Cell(y, 18).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 18).Value = row[17]; //採購單號
                    wsheet_IPS.Cell(y, 19).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 19).Value = row[18]; //採購序號
                    wsheet_IPS.Cell(y, 20).Style.NumberFormat.Format = "@";
                    wsheet_IPS.Cell(y, 20).Value = row[19]; //庫別

                    num_IPS = row[1].ToString().Trim();

                    if ((rows_count_IPS - 1) == dt_IPS.Rows.IndexOf(row)) //資料列結尾運算
                    {
                        y++;
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                        wsheet_IPS.Cell(y, 2).Style.NumberFormat.Format = "@";
                        wsheet_IPS.Cell(y, 2).Value = num_IPS;
                        wsheet_IPS.Cell(y, 3).Value = "小計";
                        wsheet_IPS.Cell(y, 9).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_IPS.Cell(y, 9).FormulaA1 = "=sum(I" + x + ":I" + (y - 1) + ")";
                        wsheet_IPS.Cell(y, 12).Style.NumberFormat.Format = "#,##0";
                        wsheet_IPS.Cell(y, 12).FormulaA1 = "=sum(L" + x + ":L" + (y - 1) + ")";

                        x = y + 1;
                        y++;

                        wsheet_IPS.Cell(y, 2).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
                        wsheet_IPS.Range("A" + y + ":T" + y).Style.Fill.BackgroundColor = XLColor.Honeydew;
                        wsheet_IPS.Cell(y, 3).Value = "總計";
                        wsheet_IPS.Cell(y, 9).Style.NumberFormat.Format = "#,##0.00";
                        wsheet_IPS.Cell(y, 9).FormulaA1 = "=SUMIF(C:C,\"小計\",I:I)";
                        wsheet_IPS.Cell(y, 12).Style.NumberFormat.Format = "#,##0";
                        wsheet_IPS.Cell(y, 12).FormulaA1 = "=SUMIF(C:C,\"小計\",L:L)";
                    }
                    y++;
                }

                //自动调整列的宽度
                wb_IPS.Worksheet(1).Columns().AdjustToContents();

                //保存文件
                wb_IPS.SaveAs(path + "\\" + fileNameWithExtension);
            }

            //打开文件
            System.Diagnostics.Process.Start(path + "\\" + fileNameWithExtension);
        }

        
       /*
        private void ExportExcelSum(DataGridView dgvDataInfo, string path, string fileNameWithExtension)
        {
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

                    }
                    j++;
                }
            }
        }*/
       
    }
}
