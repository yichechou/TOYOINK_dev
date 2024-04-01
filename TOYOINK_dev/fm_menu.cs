using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TOYOINK_dev
{
    /*
  * 使用到的TABLE
  * fm_AUOPlannedOrder: 有使用[SAP].[dbo].fm_COPTC_log、[SAP].[dbo].fm_COPTD_log，同ERP欄位 前面加入 DEL_DATE
  * 20240314 生管 梁姿儂提出 fm_AUOPlannedOrder 前回差報表的部分，調整成依照FAB,ERPNO,MATERIAL_TYPE 排序
  * 20240401 生管 梁姿儂提出 fm_AUOPlannedOrder 上傳ERP排序，加入[ERPNO]，[FAB],[ERPNO],[MATERIAL_TYPE]]
  * 
  * 
  */
    public partial class fm_menu : Form
    {
        TOYOINK_dev.Premium fm_Premium = new TOYOINK_dev.Premium();
        TOYOINK_dev.fm_Trademark fm_Trademark = new TOYOINK_dev.fm_Trademark();
        TOYOINK_dev.fm_Package7b fm_Package7b = new TOYOINK_dev.fm_Package7b();
        TOYOINK_dev.fm_AUOPlannedOrder fm_AUOPlannedOrder = new TOYOINK_dev.fm_AUOPlannedOrder();
        TOYOINK_dev.fm_AUOCOPTC fm_AUOCOPTC = new TOYOINK_dev.fm_AUOCOPTC();
        TOYOINK_dev.fm_login fm_login = new TOYOINK_dev.fm_login();
        TOYOINK_dev.fm_Acc_5b fm_Acc_5b = new TOYOINK_dev.fm_Acc_5b();
        TOYOINK_dev.fm_Acc_F22_1 fm_Acc_F22_1 = new TOYOINK_dev.fm_Acc_F22_1();
        TOYOINK_dev.fm_Acc_RelatedVOU fm_Acc_RelatedVOU = new TOYOINK_dev.fm_Acc_RelatedVOU();
        //TOYOINK_dev.fm_trycode fm_trycode = new TOYOINK_dev.fm_trycode();
        //TOYOINK_dev.fm_PC_PURTC fm_PC_PURTC = new TOYOINK_dev.fm_PC_PURTC();


        public fm_menu()
        {
            InitializeComponent();
        }

        private void fm_menu_Load(object sender, EventArgs e)
        {
            //tabControl1.SelectedIndex = 1;
        }

        private void fm_menu_FormClosed(object sender, FormClosedEventArgs e)
        {
            //Environment.Exit(Environment.ExitCode);
        }

        //接收form1資料，並顯示
        public string loginName = "", LoginFormName = "";
        //public System.Windows.Forms.Form LoginFormName;
        public int CheckForm ;
        public void show_fmlogin_loginName(string data_loginName)
        {
            loginName = data_loginName;
        }
        public void show_fmlogin_CheckForm(int data_CheckForm)
        {
            CheckForm = data_CheckForm;
        }
        public void show_fmlogin_FormName(string data_LoginFormName)
        {
            LoginFormName = data_LoginFormName;
        }

        private void btn_premium_Click(object sender, EventArgs e)
        {
            fm_Premium.Show();
            this.Hide();
        }

        private void btn__Click(object sender, EventArgs e)
        {
            fm_Trademark.Show();
            this.Hide();
        }

        private void btn_Proc_premium_Click(object sender, EventArgs e)
        {
            fm_Premium.Show();
            this.Hide();
        }

        private void btn_Package7b_Click(object sender, EventArgs e)
        {
            fm_Package7b.Show();
            this.Hide();
        }

        private void btn_AUOPlannedOrder_Click(object sender, EventArgs e)
        {
            this.Hide(); //隱藏父視窗
            fm_login.show_fmlogin_FormName("fm_AUOPlannedOrder");
            fm_login.Show();

            ////fm_AUOPlannedOrder.Show();
            //this.Hide(); //隱藏父視窗


            //fm_AUOPlannedOrder fm_AUOPlannedOrder = new fm_AUOPlannedOrder(); //創建子視窗

            //switch (fm_AUOPlannedOrder.ShowDialog(this))
            //{
            //    case DialogResult.Yes: //Form2中按下ToForm1按鈕
            //        this.Show(); //顯示父視窗
            //        this.fm_menu_Load(null, null);
            //        break;
            //    case DialogResult.No: //Form2中按下關閉鈕
            //        this.Close();  //關閉父視窗 (同時結束應用程式)
            //        break;
            //    default:
            //        break;
            //}
        }

        private void btn_AUOPlannedOrderInput_Click(object sender, EventArgs e)
        {
            //    fm_AUOPlannedOrderInput fm_AUOPlannedOrderInput = new fm_AUOPlannedOrderInput(); //創建子視窗
            //    this.Hide(); //隱藏父視窗
            //    //fm_login.show_fmlogin_FormName("fm_AUOPlannedOrderInput");
            //    //fm_login.Show();

            //    fm_AUOPlannedOrderInput.Show();
            //switch (fm_AUOPlannedOrderInput.ShowDialog(this))
            //{
            //    case DialogResult.Yes: //Form2中按下ToForm1按鈕
            //        this.Show(); //顯示父視窗
            //        this.fm_menu_Load(null, null);
            //        break;
            //    case DialogResult.No: //Form2中按下關閉鈕
            //        this.Close();  //關閉父視窗 (同時結束應用程式)
            //        break;
            //    default:
            //        break;
            //}

        }

        private void btn_AUOPlannedOrderUpdate_Click(object sender, EventArgs e)
        {
            //this.Hide(); //隱藏父視窗

            //fm_AUOPlannedOrderUpdate fm_AUOPlannedOrderUpdate = new fm_AUOPlannedOrderUpdate(); //創建子視窗

            //switch (fm_AUOPlannedOrderUpdate.ShowDialog(this))
            //{
            //    case DialogResult.Yes: //Form2中按下ToForm1按鈕
            //        this.Show(); //顯示父視窗
            //        this.fm_menu_Load(null, null);
            //        break;
            //    case DialogResult.No: //Form2中按下關閉鈕
            //        this.Close();  //關閉父視窗 (同時結束應用程式)
            //        break;
            //    default:
            //        break;
            //}
        }

        private void btn_Package5a_Click(object sender, EventArgs e)
        {
            this.Hide(); //隱藏父視窗

            fm_Package5a8a fm_Package5a8a = new fm_Package5a8a(); //創建子視窗

            switch (fm_Package5a8a.ShowDialog(this))
            {
                case DialogResult.Yes: //Form2中按下ToForm1按鈕
                    this.Show(); //顯示父視窗
                    this.fm_menu_Load(null, null);
                    break;
                case DialogResult.No: //Form2中按下關閉鈕
                    this.Close();  //關閉父視窗 (同時結束應用程式)
                    break;
                default:
                    break;
            }
        }

        private void btn_Acc_5b_Click(object sender, EventArgs e)
        {
            this.Hide(); //隱藏父視窗

            fm_Acc_5b fm_Acc_5b = new fm_Acc_5b(); //創建子視窗

            switch (fm_Acc_5b.ShowDialog(this))
            {
                case DialogResult.Yes: //Form2中按下ToForm1按鈕
                    this.Show(); //顯示父視窗
                    this.fm_menu_Load(null, null);
                    break;
                case DialogResult.No: //Form2中按下關閉鈕
                    this.Close();  //關閉父視窗 (同時結束應用程式)
                    break;
                default:
                    break;
            }
        }


        private void btn_Acc_F22_1_Click(object sender, EventArgs e)
        {
            this.Hide(); //隱藏父視窗

            fm_Acc_F22_1 fm_Acc_F22_1 = new fm_Acc_F22_1(); //創建子視窗

            switch (fm_Acc_F22_1.ShowDialog(this))
            {
                case DialogResult.Yes: //Form2中按下ToForm1按鈕
                    this.Show(); //顯示父視窗
                    this.fm_menu_Load(null, null);
                    break;
                case DialogResult.No: //Form2中按下關閉鈕
                    this.Close();  //關閉父視窗 (同時結束應用程式)
                    break;
                default:
                    break;
            }
        }

        private void btn_Acc_RelatedVOU_Click(object sender, EventArgs e)
        {
            this.Hide(); //隱藏父視窗
            fm_login.show_fmlogin_FormName("fm_Acc_RelatedVOU");
            fm_login.Show();
        }

        private void btn_AUO_NF_COPTC_Click(object sender, EventArgs e)
        {
            this.Hide(); //隱藏父視窗
            fm_login.show_fmlogin_FormName("fm_AUO_NF_COPTC");
            fm_login.Show();
        }

        private void btn_PC_PURTC_Click(object sender, EventArgs e)
        {
            this.Hide(); //隱藏父視窗
            fm_login.show_fmlogin_FormName("fm_PC_PURTC");
            fm_login.Show();

            this.Hide(); //隱藏父視窗

        }

        //fm_AUO_NF_COPTC
        private void btn_AUOCOPTC_Click(object sender, EventArgs e)
        {
            this.Hide(); //隱藏父視窗
            fm_login.show_fmlogin_FormName("fm_AUOCOPTC");
            fm_login.Show();

            //if (CheckForm == 1)
            //{
            //    fm_AUOCOPTC fm_AUOCOPTC = new fm_AUOCOPTC(); //創建子視窗

            //    switch (fm_AUOCOPTC.ShowDialog(this))
            //    {
            //        case DialogResult.Yes: //Form2中按下ToForm1按鈕
            //            this.Show(); //顯示父視窗
            //            this.fm_menu_Load(null, null);
            //            break;
            //        case DialogResult.No: //Form2中按下關閉鈕
            //            this.Close();  //關閉父視窗 (同時結束應用程式)
            //            break;
            //        default:
            //            break;
            //    }
            //}
        }

        //private void FormCloseSwitch(Form fm)
        //{
        //    this.Hide(); //隱藏父視窗

        //    Form fm = new Form.fm(); //創建子視窗

        //    switch (fm_Package5a.ShowDialog(this))
        //    {
        //        case DialogResult.Yes: //Form2中按下ToForm1按鈕
        //            this.Show(); //顯示父視窗
        //            this.fm_menu_Load(null, null);
        //            break;
        //        case DialogResult.No: //Form2中按下關閉鈕
        //            this.Close();  //關閉父視窗 (同時結束應用程式)
        //            break;
        //        default:
        //            break;
        //    }
        //}
    }
}
