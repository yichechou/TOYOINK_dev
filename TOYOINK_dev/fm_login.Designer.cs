namespace TOYOINK_dev
{
    public class Global
    {
        public static string str_erpid;
        public static string str_erpgp;
    }
    partial class fm_login
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(fm_login));
            this.btn_signin = new System.Windows.Forms.Button();
            this.btn_signout = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_id = new System.Windows.Forms.TextBox();
            this.txt_pw = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.sqlDataAdapter1 = new System.Data.SqlClient.SqlDataAdapter();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlConnection1 = new System.Data.SqlClient.SqlConnection();
            this.sqlCommand1 = new System.Data.SqlClient.SqlCommand();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_signin
            // 
            this.btn_signin.Location = new System.Drawing.Point(38, 136);
            this.btn_signin.Name = "btn_signin";
            this.btn_signin.Size = new System.Drawing.Size(107, 49);
            this.btn_signin.TabIndex = 3;
            this.btn_signin.Text = "登入";
            this.btn_signin.UseVisualStyleBackColor = true;
            this.btn_signin.Click += new System.EventHandler(this.btn_signin_Click);
            // 
            // btn_signout
            // 
            this.btn_signout.Location = new System.Drawing.Point(169, 136);
            this.btn_signout.Name = "btn_signout";
            this.btn_signout.Size = new System.Drawing.Size(107, 49);
            this.btn_signout.TabIndex = 4;
            this.btn_signout.Text = "取消";
            this.btn_signout.UseVisualStyleBackColor = true;
            this.btn_signout.Click += new System.EventHandler(this.btn_signout_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 25);
            this.label1.TabIndex = 3;
            this.label1.Text = "登入帳號";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 25);
            this.label2.TabIndex = 4;
            this.label2.Text = "登入密碼";
            // 
            // txt_id
            // 
            this.txt_id.AcceptsTab = true;
            this.txt_id.Location = new System.Drawing.Point(115, 18);
            this.txt_id.Name = "txt_id";
            this.txt_id.Size = new System.Drawing.Size(136, 34);
            this.txt_id.TabIndex = 0;
            this.txt_id.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txt_id_KeyDown);
            // 
            // txt_pw
            // 
            this.txt_pw.AcceptsTab = true;
            this.txt_pw.Location = new System.Drawing.Point(115, 60);
            this.txt_pw.Name = "txt_pw";
            this.txt_pw.PasswordChar = '*';
            this.txt_pw.Size = new System.Drawing.Size(136, 34);
            this.txt_pw.TabIndex = 1;
            this.txt_pw.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txt_pw_KeyDown);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.txt_pw);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txt_id);
            this.panel1.Location = new System.Drawing.Point(19, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(272, 115);
            this.panel1.TabIndex = 7;
            // 
            // sqlDataAdapter1
            // 
            this.sqlDataAdapter1.InsertCommand = this.sqlInsertCommand1;
            this.sqlDataAdapter1.SelectCommand = this.sqlCommand1;
            this.sqlDataAdapter1.TableMappings.AddRange(new System.Data.Common.DataTableMapping[] {
            new System.Data.Common.DataTableMapping("Table", "CFIPO", new System.Data.Common.DataColumnMapping[] {
                        new System.Data.Common.DataColumnMapping("線別", "線別"),
                        new System.Data.Common.DataColumnMapping("Sample", "Sample"),
                        new System.Data.Common.DataColumnMapping("Number", "Number"),
                        new System.Data.Common.DataColumnMapping("Item", "Item"),
                        new System.Data.Common.DataColumnMapping("Item Description", "Item Description"),
                        new System.Data.Common.DataColumnMapping("UOM", "UOM"),
                        new System.Data.Common.DataColumnMapping("Shipment Amount", "Shipment Amount"),
                        new System.Data.Common.DataColumnMapping("Quantity", "Quantity"),
                        new System.Data.Common.DataColumnMapping("Supplier", "Supplier"),
                        new System.Data.Common.DataColumnMapping("Supplier Site", "Supplier Site"),
                        new System.Data.Common.DataColumnMapping("Currency", "Currency"),
                        new System.Data.Common.DataColumnMapping("Buyer", "Buyer"),
                        new System.Data.Common.DataColumnMapping("Needdate", "Needdate"),
                        new System.Data.Common.DataColumnMapping("備註", "備註"),
                        new System.Data.Common.DataColumnMapping("序號", "序號")})});
            // 
            // sqlInsertCommand1
            // 
            this.sqlInsertCommand1.CommandText = resources.GetString("sqlInsertCommand1.CommandText");
            this.sqlInsertCommand1.Connection = this.sqlConnection1;
            this.sqlInsertCommand1.Parameters.AddRange(new System.Data.SqlClient.SqlParameter[] {
            new System.Data.SqlClient.SqlParameter("@線別", System.Data.SqlDbType.NVarChar, 0, "線別"),
            new System.Data.SqlClient.SqlParameter("@Sample", System.Data.SqlDbType.NVarChar, 0, "Sample"),
            new System.Data.SqlClient.SqlParameter("@Number", System.Data.SqlDbType.NVarChar, 0, "Number"),
            new System.Data.SqlClient.SqlParameter("@Item", System.Data.SqlDbType.NVarChar, 0, "Item"),
            new System.Data.SqlClient.SqlParameter("@Item_Description", System.Data.SqlDbType.NVarChar, 0, "Item Description"),
            new System.Data.SqlClient.SqlParameter("@UOM", System.Data.SqlDbType.NVarChar, 0, "UOM"),
            new System.Data.SqlClient.SqlParameter("@Shipment_Amount", System.Data.SqlDbType.Decimal, 0, System.Data.ParameterDirection.Input, false, ((byte)(18)), ((byte)(2)), "Shipment Amount", System.Data.DataRowVersion.Current, null),
            new System.Data.SqlClient.SqlParameter("@Quantity", System.Data.SqlDbType.Decimal, 0, System.Data.ParameterDirection.Input, false, ((byte)(18)), ((byte)(2)), "Quantity", System.Data.DataRowVersion.Current, null),
            new System.Data.SqlClient.SqlParameter("@Supplier", System.Data.SqlDbType.NVarChar, 0, "Supplier"),
            new System.Data.SqlClient.SqlParameter("@Supplier_Site", System.Data.SqlDbType.NVarChar, 0, "Supplier Site"),
            new System.Data.SqlClient.SqlParameter("@Currency", System.Data.SqlDbType.NVarChar, 0, "Currency"),
            new System.Data.SqlClient.SqlParameter("@Buyer", System.Data.SqlDbType.NVarChar, 0, "Buyer"),
            new System.Data.SqlClient.SqlParameter("@Needdate", System.Data.SqlDbType.NVarChar, 0, "Needdate"),
            new System.Data.SqlClient.SqlParameter("@備註", System.Data.SqlDbType.NVarChar, 0, "備註"),
            new System.Data.SqlClient.SqlParameter("@序號", System.Data.SqlDbType.NVarChar, 0, "序號")});
            // 
            // sqlConnection1
            // 
            //this.sqlConnection1.ConnectionString = "Data Source=192.168.128.219;Initial Catalog=A01B;User ID=pwuser;Password=sqlmis00" +
            //"3";
            this.sqlConnection1.FireInfoMessageEventOnUserErrors = false;
            // 
            // sqlCommand1
            // 
            this.sqlCommand1.CommandText = "SELECT  *\r\n  FROM YCMSMV";
            this.sqlCommand1.Connection = this.sqlConnection1;
            // 
            // fm_login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(310, 202);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btn_signout);
            this.Controls.Add(this.btn_signin);
            this.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "fm_login";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "線上系統登錄";
            this.Load += new System.EventHandler(this.fm_login_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btn_signin;
        private System.Windows.Forms.Button btn_signout;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_id;
        private System.Windows.Forms.TextBox txt_pw;
        private System.Windows.Forms.Panel panel1;
        public System.Data.SqlClient.SqlDataAdapter sqlDataAdapter1;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        public System.Data.SqlClient.SqlConnection sqlConnection1;
        private System.Data.SqlClient.SqlCommand sqlCommand1;
    }
}

