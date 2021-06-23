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

namespace TOYOINK_dev
{
    //*20210303 加入立沖登入驗證
    public partial class fm_login : Form
    {
        public MyClass MyCode;
        public string loginid = "";
        public string loginName = "" , LoginFormName = "",loginDep = "";
        //public System.Windows.Forms.Form LoginFormName;
        //string str_warning = "";
        string str_sql = "";
        //    string str_enter = ((char)13).ToString() + ((char)10).ToString();
        //    string str;

        public fm_login()
        {
            InitializeComponent();
            MyCode = new Myclass.MyClass();
            MyCode.strDbCon = "packet size=4096;user id=pwuser;password=sqlmis003;data source=192.168.128.219;persist security info=False;initial catalog=A01B;";

        }
        public void show_fmlogin_FormName(string data_LoginFormName)
        {
            LoginFormName = data_LoginFormName;
        }

        private void fm_login_Load(object sender, EventArgs e)
        {
            txt_id.Select();
            txt_id.Text = "1901329";
            txt_pw.Text = "asdf673690";

        }

        //TOYOINK_dev.fm_menu fm_menu = new TOYOINK_dev.fm_menu();

        private void btn_signin_Click(object sender, EventArgs e)
        {
            
            loginid = txt_id.Text.ToString().Trim();

            if (txt_id.Text.Trim().ToString() == "")
            {
                MessageBox.Show("帳號空值，請重新輸入", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            DataTable dt_login = new DataTable();
            //this.sqlDataAdapter1.SelectCommand.CommandText = "select NAME, MV002 as account, MV003 as password from YCMSMV " +
            //          "where MV002 ='" + loginid + "'";
            //this.sqlDataAdapter1.Fill(dt_login);
            str_sql = "select NAME, MV002 as account, MV003 as password,MV100 as dep from YCMSMV " +
                      "where MV002 ='" + loginid + "'";
            MyCode.Sql_dt(str_sql, dt_login);

            if (dt_login.Rows.Count > 0 )
            {
                if (dt_login.Rows[0]["password"].ToString().Trim() != txt_pw.Text.ToString().Trim())
                {
                    MessageBox.Show("密碼錯誤，請重新輸入", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                else
                {
                    loginName = dt_login.Rows[0]["NAME"].ToString().Trim();
                    loginDep = dt_login.Rows[0]["dep"].ToString().Trim();
                    //fm_menu.show_form1_data(loginName);
                    //fm_COP fm_COP = new fm_COP();
                    //fm_menu.Show();
                    //txt_id.Select();
                    this.Hide();

                    switch (LoginFormName) 
                    {
                        case "fm_AUOCOPTC":
                            fm_AUOCOPTC fm_AUOCOPTC = new fm_AUOCOPTC();
                            fm_AUOCOPTC.show_fmlogin_loginName(loginName);
                            fm_AUOCOPTC.show_fmlogin_CheckForm(1);
                            fm_AUOCOPTC.Show();
                            break;
                        case "fm_AUOPlannedOrder":
                            fm_AUOPlannedOrder fm_AUOPlannedOrder = new fm_AUOPlannedOrder();
                            fm_AUOPlannedOrder.show_fmlogin_loginName(loginName);
                            fm_AUOPlannedOrder.show_fmlogin_CheckForm(1);
                            fm_AUOPlannedOrder.Show();
                            break;
                        case "fm_Acc_RelatedVOU":
                            if (loginDep == "15000" || loginDep == "302" || loginDep == "300")
                            {
                                fm_Acc_RelatedVOU fm_Acc_RelatedVOU = new fm_Acc_RelatedVOU();
                                fm_Acc_RelatedVOU.show_fmlogin_loginName(loginName);
                                fm_Acc_RelatedVOU.show_fmlogin_CheckForm(1);
                                fm_Acc_RelatedVOU.Show();
                            }
                            else 
                            {
                                MessageBox.Show("非財務人員禁止使用", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                this.Show();
                                return;
                            }
                            break;
                        case "fm_AUO_NF_COPTC":
                            fm_AUO_NF_COPTC fm_AUO_NF_COPTC = new fm_AUO_NF_COPTC();
                            fm_AUO_NF_COPTC.show_fmlogin_loginName(loginName);
                            fm_AUO_NF_COPTC.show_fmlogin_CheckForm(1);
                            fm_AUO_NF_COPTC.Show();
                            break;

                        default:
                            break;
                    }

                }
            }
            else
            {
                MessageBox.Show("帳號錯誤，找不到帳號，請重新輸入", "警告訊息", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            //str_sql = "select MV002 as account, MV003 as password from S2008X64.A01B.dbo.YCMSMV " + 
            //          "where MV002 ='" + txt_id.Text.Trim().ToString() + "' and '" + txt_pw.Text.Trim().ToString() + "'";

            ////TODO: 
            //string str_sql =
            //    "select COLUMN_NAME,DATA_TYPE,IS_NULLABLE" + str_enter +
            //    "from INFORMATION_SCHEMA.COLUMNS" + str_enter +
            //    "where TABLE_NAME='CFIPO'";
            //this.sqlDataAdapter1.SelectCommand.CommandText = str_sql;
            //this.sqlDataAdapter1.Fill(dt_login);

            //fm_COP fm_COP = new fm_COP();
            //fm_COP.Show();
            //this.Hide();

            //var str = "2282400";
            ////var md5 = this.ToMD5(str);   //81DC9BDB52D04DC20036DBD8313ED055
            //label3.Text = this.ToMD5(str);
        }

        private void btn_signout_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void txt_id_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txt_pw.Focus();
                //btn_signin.TabStop = true;
            }
        }

        private void txt_pw_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btn_signin.Focus();
            }
        }

        bool IsToForm1 = false; //紀錄是否要回到Form1
        protected override void OnClosing(CancelEventArgs e) //在視窗關閉時觸發
        {
            //DialogResult dr = MessageBox.Show("\"是\"回到主畫面 \r\n \"否\"關閉程式", "是否要關閉程式"
            //    , MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            //if (dr == DialogResult.Yes)
            //{
            //    IsToForm1 = true;
            //}

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

        //private void fm_login_Activated(object sender, EventArgs e)
        //{
        //    txt_id.Focus();
        //}

        //public string ToMD5(string str)
        //{
        //    using (var cryptoMD5 = System.Security.Cryptography.MD5.Create())
        //    {
        //        //將字串編碼成 UTF8 位元組陣列
        //        var bytes = Encoding.UTF8.GetBytes(str);

        //        //取得雜湊值位元組陣列
        //        var hash = cryptoMD5.ComputeHash(bytes);

        //        //取得 MD5
        //        var md5 = BitConverter.ToString(hash)
        //            .Replace("-", String.Empty)
        //            .ToUpper();

        //        return md5;
        //    }
        //}

    }
}
