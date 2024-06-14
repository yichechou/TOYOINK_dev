namespace TOYOINK_dev
{
    partial class fm_Acc_PurSale
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(fm_Acc_PurSale));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tab_Search = new System.Windows.Forms.TabPage();
            this.dgv_Search = new System.Windows.Forms.DataGridView();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btn_up = new System.Windows.Forms.Button();
            this.txterr = new System.Windows.Forms.TextBox();
            this.btn_down = new System.Windows.Forms.Button();
            this.txt_path = new System.Windows.Forms.TextBox();
            this.btn_fileopen = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lab_status = new System.Windows.Forms.Label();
            this.btn_file = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_ToExcel = new System.Windows.Forms.Button();
            this.txt_date_e = new System.Windows.Forms.TextBox();
            this.btn_search = new System.Windows.Forms.Button();
            this.Btn_date_e = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txt_date_s = new System.Windows.Forms.TextBox();
            this.Btn_date_s = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tab_Search.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Search)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tab_Search);
            this.tabControl1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tabControl1.HotTrack = true;
            this.tabControl1.Location = new System.Drawing.Point(12, 195);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1285, 645);
            this.tabControl1.TabIndex = 11;
            // 
            // tab_Search
            // 
            this.tab_Search.Controls.Add(this.dgv_Search);
            this.tab_Search.Location = new System.Drawing.Point(4, 34);
            this.tab_Search.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tab_Search.Name = "tab_Search";
            this.tab_Search.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tab_Search.Size = new System.Drawing.Size(1277, 607);
            this.tab_Search.TabIndex = 1;
            this.tab_Search.Text = "查詢結果";
            this.tab_Search.UseVisualStyleBackColor = true;
            // 
            // dgv_Search
            // 
            this.dgv_Search.AllowUserToAddRows = false;
            this.dgv_Search.AllowUserToDeleteRows = false;
            this.dgv_Search.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_Search.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_Search.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_Search.Location = new System.Drawing.Point(4, 5);
            this.dgv_Search.Name = "dgv_Search";
            this.dgv_Search.ReadOnly = true;
            this.dgv_Search.RowHeadersWidth = 51;
            this.dgv_Search.RowTemplate.Height = 27;
            this.dgv_Search.Size = new System.Drawing.Size(1269, 597);
            this.dgv_Search.TabIndex = 0;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.tabControl1);
            this.panel2.Controls.Add(this.btn_up);
            this.panel2.Controls.Add(this.txterr);
            this.panel2.Controls.Add(this.btn_down);
            this.panel2.Controls.Add(this.txt_path);
            this.panel2.Controls.Add(this.btn_fileopen);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.lab_status);
            this.panel2.Controls.Add(this.btn_file);
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Controls.Add(this.txt_date_e);
            this.panel2.Controls.Add(this.btn_search);
            this.panel2.Controls.Add(this.Btn_date_e);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.txt_date_s);
            this.panel2.Controls.Add(this.Btn_date_s);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1309, 852);
            this.panel2.TabIndex = 10;
            // 
            // btn_up
            // 
            this.btn_up.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_up.Location = new System.Drawing.Point(502, 19);
            this.btn_up.Name = "btn_up";
            this.btn_up.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btn_up.Size = new System.Drawing.Size(38, 35);
            this.btn_up.TabIndex = 55;
            this.btn_up.Text = "▶";
            this.btn_up.UseVisualStyleBackColor = true;
            this.btn_up.Click += new System.EventHandler(this.btn_up_Click);
            // 
            // txterr
            // 
            this.txterr.Font = new System.Drawing.Font("微軟正黑體", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txterr.Location = new System.Drawing.Point(826, 13);
            this.txterr.Multiline = true;
            this.txterr.Name = "txterr";
            this.txterr.ReadOnly = true;
            this.txterr.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txterr.Size = new System.Drawing.Size(469, 176);
            this.txterr.TabIndex = 22;
            // 
            // btn_down
            // 
            this.btn_down.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_down.Location = new System.Drawing.Point(453, 19);
            this.btn_down.Name = "btn_down";
            this.btn_down.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btn_down.Size = new System.Drawing.Size(38, 35);
            this.btn_down.TabIndex = 54;
            this.btn_down.Text = "◀";
            this.btn_down.UseVisualStyleBackColor = true;
            this.btn_down.Click += new System.EventHandler(this.btn_down_Click);
            // 
            // txt_path
            // 
            this.txt_path.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_path.Location = new System.Drawing.Point(112, 73);
            this.txt_path.Name = "txt_path";
            this.txt_path.ReadOnly = true;
            this.txt_path.Size = new System.Drawing.Size(433, 34);
            this.txt_path.TabIndex = 23;
            this.txt_path.Text = "D:\\";
            // 
            // btn_fileopen
            // 
            this.btn_fileopen.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_fileopen.Location = new System.Drawing.Point(693, 65);
            this.btn_fileopen.Name = "btn_fileopen";
            this.btn_fileopen.Size = new System.Drawing.Size(115, 51);
            this.btn_fileopen.TabIndex = 37;
            this.btn_fileopen.Text = "打開位置";
            this.btn_fileopen.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(14, 78);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 25);
            this.label1.TabIndex = 21;
            this.label1.Text = "存檔位置";
            // 
            // lab_status
            // 
            this.lab_status.BackColor = System.Drawing.SystemColors.Info;
            this.lab_status.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_status.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lab_status.Location = new System.Drawing.Point(556, 13);
            this.lab_status.Name = "lab_status";
            this.lab_status.Size = new System.Drawing.Size(252, 45);
            this.lab_status.TabIndex = 30;
            this.lab_status.Text = " 請先選擇 單據日期";
            this.lab_status.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btn_file
            // 
            this.btn_file.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_file.Location = new System.Drawing.Point(561, 65);
            this.btn_file.Name = "btn_file";
            this.btn_file.Size = new System.Drawing.Size(115, 51);
            this.btn_file.TabIndex = 2;
            this.btn_file.Text = "選擇路徑";
            this.btn_file.UseVisualStyleBackColor = true;
            this.btn_file.Click += new System.EventHandler(this.btn_file_Click);
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.panel3.Controls.Add(this.label2);
            this.panel3.Controls.Add(this.btn_ToExcel);
            this.panel3.Location = new System.Drawing.Point(19, 122);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(526, 67);
            this.panel3.TabIndex = 36;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(1, 14);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(140, 36);
            this.label2.TabIndex = 37;
            this.label2.Text = "Excel轉出";
            // 
            // btn_ToExcel
            // 
            this.btn_ToExcel.BackColor = System.Drawing.SystemColors.Control;
            this.btn_ToExcel.Enabled = false;
            this.btn_ToExcel.Font = new System.Drawing.Font("微軟正黑體", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_ToExcel.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btn_ToExcel.Location = new System.Drawing.Point(147, 7);
            this.btn_ToExcel.Name = "btn_ToExcel";
            this.btn_ToExcel.Size = new System.Drawing.Size(356, 50);
            this.btn_ToExcel.TabIndex = 25;
            this.btn_ToExcel.Text = "進貨之銷貨資料";
            this.btn_ToExcel.UseVisualStyleBackColor = false;
            this.btn_ToExcel.Click += new System.EventHandler(this.btn_ToExcel_Click);
            // 
            // txt_date_e
            // 
            this.txt_date_e.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.txt_date_e.Cursor = System.Windows.Forms.Cursors.No;
            this.txt_date_e.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_date_e.Location = new System.Drawing.Point(291, 20);
            this.txt_date_e.Name = "txt_date_e";
            this.txt_date_e.ReadOnly = true;
            this.txt_date_e.Size = new System.Drawing.Size(110, 34);
            this.txt_date_e.TabIndex = 12;
            this.txt_date_e.Text = "20231231";
            // 
            // btn_search
            // 
            this.btn_search.BackColor = System.Drawing.Color.SteelBlue;
            this.btn_search.Font = new System.Drawing.Font("微軟正黑體", 19.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_search.ForeColor = System.Drawing.Color.White;
            this.btn_search.Location = new System.Drawing.Point(561, 122);
            this.btn_search.Name = "btn_search";
            this.btn_search.Size = new System.Drawing.Size(247, 67);
            this.btn_search.TabIndex = 1;
            this.btn_search.Text = "查詢";
            this.btn_search.UseVisualStyleBackColor = false;
            this.btn_search.Click += new System.EventHandler(this.btn_search_Click);
            // 
            // Btn_date_e
            // 
            this.Btn_date_e.Image = ((System.Drawing.Image)(resources.GetObject("Btn_date_e.Image")));
            this.Btn_date_e.Location = new System.Drawing.Point(406, 20);
            this.Btn_date_e.Name = "Btn_date_e";
            this.Btn_date_e.Size = new System.Drawing.Size(35, 33);
            this.Btn_date_e.TabIndex = 13;
            this.Btn_date_e.Click += new System.EventHandler(this.Btn_date_e_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(263, 24);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(27, 25);
            this.label3.TabIndex = 33;
            this.label3.Text = "~";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(12, 25);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(92, 25);
            this.label5.TabIndex = 27;
            this.label5.Text = "日期區間";
            // 
            // txt_date_s
            // 
            this.txt_date_s.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.txt_date_s.Cursor = System.Windows.Forms.Cursors.No;
            this.txt_date_s.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_date_s.Location = new System.Drawing.Point(111, 20);
            this.txt_date_s.Name = "txt_date_s";
            this.txt_date_s.ReadOnly = true;
            this.txt_date_s.Size = new System.Drawing.Size(110, 34);
            this.txt_date_s.TabIndex = 10;
            this.txt_date_s.Text = "20230101";
            // 
            // Btn_date_s
            // 
            this.Btn_date_s.Image = ((System.Drawing.Image)(resources.GetObject("Btn_date_s.Image")));
            this.Btn_date_s.Location = new System.Drawing.Point(226, 20);
            this.Btn_date_s.Name = "Btn_date_s";
            this.Btn_date_s.Size = new System.Drawing.Size(35, 33);
            this.Btn_date_s.TabIndex = 11;
            this.Btn_date_s.Click += new System.EventHandler(this.Btn_date_s_Click);
            // 
            // fm_Acc_PurSale
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1309, 852);
            this.Controls.Add(this.panel2);
            this.Name = "fm_Acc_PurSale";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "進貨之銷售資料(20240614 0830)";
            this.Load += new System.EventHandler(this.fm_Acc_PurSale_Load);
            this.tabControl1.ResumeLayout(false);
            this.tab_Search.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Search)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tab_Search;
        private System.Windows.Forms.DataGridView dgv_Search;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btn_up;
        private System.Windows.Forms.TextBox txterr;
        private System.Windows.Forms.Button btn_down;
        private System.Windows.Forms.TextBox txt_path;
        private System.Windows.Forms.Button btn_fileopen;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lab_status;
        private System.Windows.Forms.Button btn_file;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_ToExcel;
        private System.Windows.Forms.TextBox txt_date_e;
        private System.Windows.Forms.Button btn_search;
        private System.Windows.Forms.Button Btn_date_e;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txt_date_s;
        private System.Windows.Forms.Button Btn_date_s;
    }
}