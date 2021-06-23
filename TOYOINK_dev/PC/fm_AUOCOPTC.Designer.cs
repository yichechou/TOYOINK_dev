namespace TOYOINK_dev
{
    partial class fm_AUOCOPTC
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(fm_AUOCOPTC));
            this.lab_Nowdate = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.lab_status = new System.Windows.Forms.Label();
            this.cob_建立者 = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.lab_num2 = new System.Windows.Forms.Label();
            this.lab_num1 = new System.Windows.Forms.Label();
            this.button_單據日期 = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.txterr = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txt_path = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_file = new System.Windows.Forms.Button();
            this.btn_toerp = new System.Windows.Forms.Button();
            this.textBox_單據日期 = new System.Windows.Forms.TextBox();
            this.btn_erpup = new System.Windows.Forms.Button();
            this.sqlDataAdapter1 = new System.Data.SqlClient.SqlDataAdapter();
            this.sqlInsertCommand1 = new System.Data.SqlClient.SqlCommand();
            this.sqlConnection1 = new System.Data.SqlClient.SqlConnection();
            this.sqlCommand1 = new System.Data.SqlClient.SqlCommand();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.dgv_excel = new System.Windows.Forms.DataGridView();
            this.tabCtl_data = new System.Windows.Forms.TabControl();
            this.tabPage_excel = new System.Windows.Forms.TabPage();
            this.tabPage_CFIPO = new System.Windows.Forms.TabPage();
            this.dgv_cfipo = new System.Windows.Forms.DataGridView();
            this.tabPage_tc = new System.Windows.Forms.TabPage();
            this.dgv_tc = new System.Windows.Forms.DataGridView();
            this.tabPage_td = new System.Windows.Forms.TabPage();
            this.dgv_td = new System.Windows.Forms.DataGridView();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_excel)).BeginInit();
            this.tabCtl_data.SuspendLayout();
            this.tabPage_excel.SuspendLayout();
            this.tabPage_CFIPO.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_cfipo)).BeginInit();
            this.tabPage_tc.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_tc)).BeginInit();
            this.tabPage_td.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_td)).BeginInit();
            this.SuspendLayout();
            // 
            // lab_Nowdate
            // 
            this.lab_Nowdate.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lab_Nowdate.BackColor = System.Drawing.Color.Gainsboro;
            this.lab_Nowdate.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_Nowdate.Location = new System.Drawing.Point(349, 99);
            this.lab_Nowdate.Name = "lab_Nowdate";
            this.lab_Nowdate.Size = new System.Drawing.Size(112, 33);
            this.lab_Nowdate.TabIndex = 26;
            this.lab_Nowdate.Text = "20190101";
            this.lab_Nowdate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label6.Location = new System.Drawing.Point(256, 102);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(92, 25);
            this.label6.TabIndex = 25;
            this.label6.Text = "建立日期";
            // 
            // lab_status
            // 
            this.lab_status.BackColor = System.Drawing.SystemColors.Info;
            this.lab_status.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_status.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lab_status.Location = new System.Drawing.Point(478, 99);
            this.lab_status.Name = "lab_status";
            this.lab_status.Size = new System.Drawing.Size(285, 45);
            this.lab_status.TabIndex = 20;
            this.lab_status.Text = " 請先選擇 單據日期";
            this.lab_status.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cob_建立者
            // 
            this.cob_建立者.Enabled = false;
            this.cob_建立者.FormattingEnabled = true;
            this.cob_建立者.Location = new System.Drawing.Point(96, 99);
            this.cob_建立者.Name = "cob_建立者";
            this.cob_建立者.Size = new System.Drawing.Size(154, 37);
            this.cob_建立者.TabIndex = 24;
            this.cob_建立者.SelectedIndexChanged += new System.EventHandler(this.cob_建立者_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(4, 102);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 25);
            this.label2.TabIndex = 23;
            this.label2.Text = "建  立  者";
            // 
            // lab_num2
            // 
            this.lab_num2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lab_num2.BackColor = System.Drawing.Color.Turquoise;
            this.lab_num2.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_num2.Location = new System.Drawing.Point(415, 12);
            this.lab_num2.Name = "lab_num2";
            this.lab_num2.Size = new System.Drawing.Size(96, 32);
            this.lab_num2.TabIndex = 22;
            this.lab_num2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lab_num1
            // 
            this.lab_num1.BackColor = System.Drawing.Color.Turquoise;
            this.lab_num1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_num1.Location = new System.Drawing.Point(307, 12);
            this.lab_num1.Name = "lab_num1";
            this.lab_num1.Size = new System.Drawing.Size(49, 31);
            this.lab_num1.TabIndex = 21;
            this.lab_num1.Text = "220";
            this.lab_num1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button_單據日期
            // 
            this.button_單據日期.Image = ((System.Drawing.Image)(resources.GetObject("button_單據日期.Image")));
            this.button_單據日期.Location = new System.Drawing.Point(216, 10);
            this.button_單據日期.Name = "button_單據日期";
            this.button_單據日期.Size = new System.Drawing.Size(35, 33);
            this.button_單據日期.TabIndex = 18;
            this.button_單據日期.Click += new System.EventHandler(this.button_單據日期_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(5, 15);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(92, 25);
            this.label5.TabIndex = 15;
            this.label5.Text = "單據日期";
            // 
            // txterr
            // 
            this.txterr.Font = new System.Drawing.Font("微軟正黑體", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txterr.Location = new System.Drawing.Point(3, 0);
            this.txterr.Multiline = true;
            this.txterr.Name = "txterr";
            this.txterr.ReadOnly = true;
            this.txterr.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txterr.Size = new System.Drawing.Size(232, 143);
            this.txterr.TabIndex = 0;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.txterr);
            this.panel2.Location = new System.Drawing.Point(923, 7);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(238, 146);
            this.panel2.TabIndex = 27;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel2);
            this.panel1.Controls.Add(this.lab_Nowdate);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.lab_status);
            this.panel1.Controls.Add(this.cob_建立者);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.lab_num2);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.txt_path);
            this.panel1.Controls.Add(this.lab_num1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.btn_file);
            this.panel1.Controls.Add(this.button_單據日期);
            this.panel1.Controls.Add(this.btn_toerp);
            this.panel1.Controls.Add(this.textBox_單據日期);
            this.panel1.Controls.Add(this.btn_erpup);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Location = new System.Drawing.Point(12, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1169, 162);
            this.panel1.TabIndex = 17;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(257, 15);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(52, 25);
            this.label3.TabIndex = 11;
            this.label3.Text = "單別";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(362, 15);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(52, 25);
            this.label4.TabIndex = 12;
            this.label4.Text = "單號";
            // 
            // txt_path
            // 
            this.txt_path.Location = new System.Drawing.Point(95, 52);
            this.txt_path.Name = "txt_path";
            this.txt_path.ReadOnly = true;
            this.txt_path.Size = new System.Drawing.Size(492, 38);
            this.txt_path.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(3, 57);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "檔案路徑";
            // 
            // btn_file
            // 
            this.btn_file.Location = new System.Drawing.Point(593, 51);
            this.btn_file.Name = "btn_file";
            this.btn_file.Size = new System.Drawing.Size(128, 37);
            this.btn_file.TabIndex = 2;
            this.btn_file.Text = "選擇檔案";
            this.btn_file.UseVisualStyleBackColor = true;
            this.btn_file.Click += new System.EventHandler(this.btn_file_Click);
            // 
            // btn_toerp
            // 
            this.btn_toerp.BackColor = System.Drawing.SystemColors.Control;
            this.btn_toerp.Enabled = false;
            this.btn_toerp.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_toerp.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btn_toerp.Location = new System.Drawing.Point(775, 6);
            this.btn_toerp.Name = "btn_toerp";
            this.btn_toerp.Size = new System.Drawing.Size(137, 72);
            this.btn_toerp.TabIndex = 7;
            this.btn_toerp.Text = "轉為ERP格式";
            this.btn_toerp.UseVisualStyleBackColor = false;
            this.btn_toerp.Click += new System.EventHandler(this.btn_toerp_Click);
            // 
            // textBox_單據日期
            // 
            this.textBox_單據日期.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.textBox_單據日期.Cursor = System.Windows.Forms.Cursors.No;
            this.textBox_單據日期.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.textBox_單據日期.Location = new System.Drawing.Point(96, 10);
            this.textBox_單據日期.Name = "textBox_單據日期";
            this.textBox_單據日期.ReadOnly = true;
            this.textBox_單據日期.Size = new System.Drawing.Size(114, 34);
            this.textBox_單據日期.TabIndex = 19;
            this.textBox_單據日期.TextChanged += new System.EventHandler(this.textBox_單據日期_TextChanged);
            // 
            // btn_erpup
            // 
            this.btn_erpup.BackColor = System.Drawing.SystemColors.Control;
            this.btn_erpup.Enabled = false;
            this.btn_erpup.Font = new System.Drawing.Font("微軟正黑體", 13.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_erpup.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btn_erpup.Location = new System.Drawing.Point(775, 82);
            this.btn_erpup.Name = "btn_erpup";
            this.btn_erpup.Size = new System.Drawing.Size(137, 72);
            this.btn_erpup.TabIndex = 6;
            this.btn_erpup.Text = "上傳至ERP";
            this.btn_erpup.UseVisualStyleBackColor = false;
            this.btn_erpup.Click += new System.EventHandler(this.btn_erpup_Click);
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
            this.sqlConnection1.ConnectionString = "Data Source=192.168.128.219;Initial Catalog=A01A;User ID=pwuser;Password=sqlmis00" +
    "3";
            this.sqlConnection1.FireInfoMessageEventOnUserErrors = false;
            // 
            // sqlCommand1
            // 
            this.sqlCommand1.CommandText = "SELECT  *\r\n  FROM CFIPO";
            this.sqlCommand1.Connection = this.sqlConnection1;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // dgv_excel
            // 
            this.dgv_excel.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_excel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_excel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_excel.Location = new System.Drawing.Point(0, 0);
            this.dgv_excel.Name = "dgv_excel";
            this.dgv_excel.ReadOnly = true;
            this.dgv_excel.RowHeadersWidth = 51;
            this.dgv_excel.RowTemplate.Height = 27;
            this.dgv_excel.Size = new System.Drawing.Size(1161, 534);
            this.dgv_excel.TabIndex = 4;
            // 
            // tabCtl_data
            // 
            this.tabCtl_data.Controls.Add(this.tabPage_excel);
            this.tabCtl_data.Controls.Add(this.tabPage_CFIPO);
            this.tabCtl_data.Controls.Add(this.tabPage_tc);
            this.tabCtl_data.Controls.Add(this.tabPage_td);
            this.tabCtl_data.Location = new System.Drawing.Point(12, 180);
            this.tabCtl_data.Name = "tabCtl_data";
            this.tabCtl_data.SelectedIndex = 0;
            this.tabCtl_data.Size = new System.Drawing.Size(1169, 576);
            this.tabCtl_data.TabIndex = 16;
            // 
            // tabPage_excel
            // 
            this.tabPage_excel.Controls.Add(this.dgv_excel);
            this.tabPage_excel.Location = new System.Drawing.Point(4, 38);
            this.tabPage_excel.Name = "tabPage_excel";
            this.tabPage_excel.Size = new System.Drawing.Size(1161, 534);
            this.tabPage_excel.TabIndex = 0;
            this.tabPage_excel.Text = "來源Excel";
            this.tabPage_excel.UseVisualStyleBackColor = true;
            // 
            // tabPage_CFIPO
            // 
            this.tabPage_CFIPO.Controls.Add(this.dgv_cfipo);
            this.tabPage_CFIPO.Location = new System.Drawing.Point(4, 38);
            this.tabPage_CFIPO.Name = "tabPage_CFIPO";
            this.tabPage_CFIPO.Size = new System.Drawing.Size(1161, 534);
            this.tabPage_CFIPO.TabIndex = 3;
            this.tabPage_CFIPO.Text = "整理結果";
            this.tabPage_CFIPO.UseVisualStyleBackColor = true;
            // 
            // dgv_cfipo
            // 
            this.dgv_cfipo.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_cfipo.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgv_cfipo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_cfipo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_cfipo.Location = new System.Drawing.Point(0, 0);
            this.dgv_cfipo.Name = "dgv_cfipo";
            this.dgv_cfipo.ReadOnly = true;
            this.dgv_cfipo.RowHeadersWidth = 51;
            this.dgv_cfipo.RowTemplate.Height = 27;
            this.dgv_cfipo.Size = new System.Drawing.Size(1161, 534);
            this.dgv_cfipo.TabIndex = 0;
            // 
            // tabPage_tc
            // 
            this.tabPage_tc.Controls.Add(this.dgv_tc);
            this.tabPage_tc.Location = new System.Drawing.Point(4, 38);
            this.tabPage_tc.Name = "tabPage_tc";
            this.tabPage_tc.Size = new System.Drawing.Size(1161, 534);
            this.tabPage_tc.TabIndex = 1;
            this.tabPage_tc.Text = "單頭 COPTC";
            this.tabPage_tc.UseVisualStyleBackColor = true;
            // 
            // dgv_tc
            // 
            this.dgv_tc.AllowUserToDeleteRows = false;
            this.dgv_tc.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_tc.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_tc.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_tc.Location = new System.Drawing.Point(0, 0);
            this.dgv_tc.Name = "dgv_tc";
            this.dgv_tc.ReadOnly = true;
            this.dgv_tc.RowHeadersWidth = 51;
            this.dgv_tc.RowTemplate.Height = 27;
            this.dgv_tc.Size = new System.Drawing.Size(1161, 534);
            this.dgv_tc.TabIndex = 4;
            // 
            // tabPage_td
            // 
            this.tabPage_td.Controls.Add(this.dgv_td);
            this.tabPage_td.Location = new System.Drawing.Point(4, 38);
            this.tabPage_td.Name = "tabPage_td";
            this.tabPage_td.Size = new System.Drawing.Size(1161, 534);
            this.tabPage_td.TabIndex = 2;
            this.tabPage_td.Text = "單身 COPTD";
            this.tabPage_td.UseVisualStyleBackColor = true;
            // 
            // dgv_td
            // 
            this.dgv_td.AllowUserToDeleteRows = false;
            this.dgv_td.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_td.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_td.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_td.Location = new System.Drawing.Point(0, 0);
            this.dgv_td.Name = "dgv_td";
            this.dgv_td.ReadOnly = true;
            this.dgv_td.RowHeadersWidth = 51;
            this.dgv_td.RowTemplate.Height = 27;
            this.dgv_td.Size = new System.Drawing.Size(1161, 534);
            this.dgv_td.TabIndex = 5;
            // 
            // fm_AUOCOPTC
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(14F, 29F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1195, 770);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.tabCtl_data);
            this.Font = new System.Drawing.Font("微軟正黑體", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.Name = "fm_AUOCOPTC";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "客戶訂單匯入 For 友達 (20210623 1050)";
            this.Load += new System.EventHandler(this.fm_AUOCOPTC_Load);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_excel)).EndInit();
            this.tabCtl_data.ResumeLayout(false);
            this.tabPage_excel.ResumeLayout(false);
            this.tabPage_CFIPO.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_cfipo)).EndInit();
            this.tabPage_tc.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_tc)).EndInit();
            this.tabPage_td.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_td)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lab_Nowdate;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label lab_status;
        private System.Windows.Forms.ComboBox cob_建立者;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lab_num2;
        private System.Windows.Forms.Label lab_num1;
        private System.Windows.Forms.Button button_單據日期;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txterr;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txt_path;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_file;
        private System.Windows.Forms.Button btn_toerp;
        private System.Windows.Forms.TextBox textBox_單據日期;
        private System.Windows.Forms.Button btn_erpup;
        public System.Data.SqlClient.SqlDataAdapter sqlDataAdapter1;
        private System.Data.SqlClient.SqlCommand sqlInsertCommand1;
        public System.Data.SqlClient.SqlConnection sqlConnection1;
        private System.Data.SqlClient.SqlCommand sqlCommand1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.DataGridView dgv_excel;
        private System.Windows.Forms.TabControl tabCtl_data;
        private System.Windows.Forms.TabPage tabPage_excel;
        private System.Windows.Forms.TabPage tabPage_CFIPO;
        private System.Windows.Forms.DataGridView dgv_cfipo;
        private System.Windows.Forms.TabPage tabPage_tc;
        private System.Windows.Forms.DataGridView dgv_tc;
        private System.Windows.Forms.TabPage tabPage_td;
        private System.Windows.Forms.DataGridView dgv_td;
    }
}