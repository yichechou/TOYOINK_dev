namespace TOYOINK_dev
{
    partial class fm_Package7b
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(fm_Package7b));
            this.btn_up = new System.Windows.Forms.Button();
            this.btn_down = new System.Windows.Forms.Button();
            this.btn_fileopen = new System.Windows.Forms.Button();
            this.lab_status = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.btn_IPS = new System.Windows.Forms.Button();
            this.btn_INV = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_7B = new System.Windows.Forms.Button();
            this.txterr = new System.Windows.Forms.TextBox();
            this.Btn_date_e = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.Btn_date_s = new System.Windows.Forms.Button();
            this.txt_date_s = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.btn_search = new System.Windows.Forms.Button();
            this.txt_date_e = new System.Windows.Forms.TextBox();
            this.btn_file = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txt_path = new System.Windows.Forms.TextBox();
            this.dgv_7B = new System.Windows.Forms.DataGridView();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tab_7B = new System.Windows.Forms.TabPage();
            this.tab_INV = new System.Windows.Forms.TabPage();
            this.dgv_INV = new System.Windows.Forms.DataGridView();
            this.tab_IPS = new System.Windows.Forms.TabPage();
            this.dgv_IPS = new System.Windows.Forms.DataGridView();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_7B)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tab_7B.SuspendLayout();
            this.tab_INV.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_INV)).BeginInit();
            this.tab_IPS.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_IPS)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_up
            // 
            this.btn_up.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_up.Location = new System.Drawing.Point(517, 19);
            this.btn_up.Name = "btn_up";
            this.btn_up.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btn_up.Size = new System.Drawing.Size(30, 35);
            this.btn_up.TabIndex = 55;
            this.btn_up.Text = "▶";
            this.btn_up.UseVisualStyleBackColor = true;
            this.btn_up.Click += new System.EventHandler(this.btn_up_Click);
            // 
            // btn_down
            // 
            this.btn_down.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_down.Location = new System.Drawing.Point(481, 19);
            this.btn_down.Name = "btn_down";
            this.btn_down.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btn_down.Size = new System.Drawing.Size(30, 35);
            this.btn_down.TabIndex = 54;
            this.btn_down.Text = "◀";
            this.btn_down.UseVisualStyleBackColor = true;
            this.btn_down.Click += new System.EventHandler(this.btn_down_Click);
            // 
            // btn_fileopen
            // 
            this.btn_fileopen.Location = new System.Drawing.Point(633, 65);
            this.btn_fileopen.Name = "btn_fileopen";
            this.btn_fileopen.Size = new System.Drawing.Size(115, 37);
            this.btn_fileopen.TabIndex = 37;
            this.btn_fileopen.Text = "打開位置";
            this.btn_fileopen.UseVisualStyleBackColor = true;
            this.btn_fileopen.Click += new System.EventHandler(this.btn_fileopen_Click);
            // 
            // lab_status
            // 
            this.lab_status.BackColor = System.Drawing.SystemColors.Info;
            this.lab_status.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lab_status.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lab_status.Location = new System.Drawing.Point(557, 13);
            this.lab_status.Name = "lab_status";
            this.lab_status.Size = new System.Drawing.Size(190, 45);
            this.lab_status.TabIndex = 30;
            this.lab_status.Text = " 請先選擇 單據日期";
            this.lab_status.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.panel3.Controls.Add(this.btn_IPS);
            this.panel3.Controls.Add(this.btn_INV);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Controls.Add(this.btn_7B);
            this.panel3.Location = new System.Drawing.Point(19, 106);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(590, 63);
            this.panel3.TabIndex = 36;
            // 
            // btn_IPS
            // 
            this.btn_IPS.BackColor = System.Drawing.SystemColors.Control;
            this.btn_IPS.Enabled = false;
            this.btn_IPS.Font = new System.Drawing.Font("微軟正黑體", 13.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_IPS.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btn_IPS.Location = new System.Drawing.Point(450, 7);
            this.btn_IPS.Name = "btn_IPS";
            this.btn_IPS.Size = new System.Drawing.Size(116, 50);
            this.btn_IPS.TabIndex = 40;
            this.btn_IPS.Text = "在途明細";
            this.btn_IPS.UseVisualStyleBackColor = false;
            this.btn_IPS.Click += new System.EventHandler(this.btn_IPS_Click);
            // 
            // btn_INV
            // 
            this.btn_INV.BackColor = System.Drawing.SystemColors.Control;
            this.btn_INV.Enabled = false;
            this.btn_INV.Font = new System.Drawing.Font("微軟正黑體", 13.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_INV.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btn_INV.Location = new System.Drawing.Point(323, 7);
            this.btn_INV.Name = "btn_INV";
            this.btn_INV.Size = new System.Drawing.Size(116, 50);
            this.btn_INV.TabIndex = 39;
            this.btn_INV.Text = "庫存明細";
            this.btn_INV.UseVisualStyleBackColor = false;
            this.btn_INV.Click += new System.EventHandler(this.btn_INV_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 16.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(7, 14);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(140, 36);
            this.label2.TabIndex = 37;
            this.label2.Text = "Excel轉出";
            // 
            // btn_7B
            // 
            this.btn_7B.BackColor = System.Drawing.SystemColors.Control;
            this.btn_7B.Enabled = false;
            this.btn_7B.Font = new System.Drawing.Font("微軟正黑體", 13.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_7B.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btn_7B.Location = new System.Drawing.Point(153, 7);
            this.btn_7B.Name = "btn_7B";
            this.btn_7B.Size = new System.Drawing.Size(159, 50);
            this.btn_7B.TabIndex = 25;
            this.btn_7B.Text = "7B彙總及明細";
            this.btn_7B.UseVisualStyleBackColor = false;
            this.btn_7B.Click += new System.EventHandler(this.btn_7B_Click);
            // 
            // txterr
            // 
            this.txterr.Font = new System.Drawing.Font("微軟正黑體", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txterr.Location = new System.Drawing.Point(758, 13);
            this.txterr.Multiline = true;
            this.txterr.Name = "txterr";
            this.txterr.ReadOnly = true;
            this.txterr.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txterr.Size = new System.Drawing.Size(373, 156);
            this.txterr.TabIndex = 22;
            // 
            // Btn_date_e
            // 
            this.Btn_date_e.Image = ((System.Drawing.Image)(resources.GetObject("Btn_date_e.Image")));
            this.Btn_date_e.Location = new System.Drawing.Point(439, 20);
            this.Btn_date_e.Name = "Btn_date_e";
            this.Btn_date_e.Size = new System.Drawing.Size(35, 33);
            this.Btn_date_e.TabIndex = 13;
            this.Btn_date_e.Click += new System.EventHandler(this.Btn_date_e_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(12, 25);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(132, 25);
            this.label5.TabIndex = 27;
            this.label5.Text = "庫存日期區間";
            // 
            // Btn_date_s
            // 
            this.Btn_date_s.Image = ((System.Drawing.Image)(resources.GetObject("Btn_date_s.Image")));
            this.Btn_date_s.Location = new System.Drawing.Point(259, 20);
            this.Btn_date_s.Name = "Btn_date_s";
            this.Btn_date_s.Size = new System.Drawing.Size(35, 33);
            this.Btn_date_s.TabIndex = 11;
            this.Btn_date_s.Click += new System.EventHandler(this.Btn_date_s_Click);
            // 
            // txt_date_s
            // 
            this.txt_date_s.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.txt_date_s.Cursor = System.Windows.Forms.Cursors.No;
            this.txt_date_s.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_date_s.Location = new System.Drawing.Point(144, 20);
            this.txt_date_s.Name = "txt_date_s";
            this.txt_date_s.ReadOnly = true;
            this.txt_date_s.Size = new System.Drawing.Size(110, 34);
            this.txt_date_s.TabIndex = 10;
            this.txt_date_s.Text = "20200101";
            this.txt_date_s.TextChanged += new System.EventHandler(this.txt_date_s_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(296, 24);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(27, 25);
            this.label3.TabIndex = 33;
            this.label3.Text = "~";
            // 
            // btn_search
            // 
            this.btn_search.BackColor = System.Drawing.Color.SteelBlue;
            this.btn_search.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btn_search.ForeColor = System.Drawing.Color.White;
            this.btn_search.Location = new System.Drawing.Point(615, 105);
            this.btn_search.Name = "btn_search";
            this.btn_search.Size = new System.Drawing.Size(135, 65);
            this.btn_search.TabIndex = 1;
            this.btn_search.Text = "查詢";
            this.btn_search.UseVisualStyleBackColor = false;
            this.btn_search.Click += new System.EventHandler(this.btn_search_Click);
            // 
            // txt_date_e
            // 
            this.txt_date_e.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            this.txt_date_e.Cursor = System.Windows.Forms.Cursors.No;
            this.txt_date_e.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt_date_e.Location = new System.Drawing.Point(324, 20);
            this.txt_date_e.Name = "txt_date_e";
            this.txt_date_e.ReadOnly = true;
            this.txt_date_e.Size = new System.Drawing.Size(110, 34);
            this.txt_date_e.TabIndex = 12;
            this.txt_date_e.Text = "20200131";
            this.txt_date_e.TextChanged += new System.EventHandler(this.txt_date_e_TextChanged);
            // 
            // btn_file
            // 
            this.btn_file.Location = new System.Drawing.Point(512, 65);
            this.btn_file.Name = "btn_file";
            this.btn_file.Size = new System.Drawing.Size(115, 37);
            this.btn_file.TabIndex = 2;
            this.btn_file.Text = "選擇路徑";
            this.btn_file.UseVisualStyleBackColor = true;
            this.btn_file.Click += new System.EventHandler(this.btn_file_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(14, 71);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(92, 25);
            this.label1.TabIndex = 21;
            this.label1.Text = "存檔位置";
            // 
            // txt_path
            // 
            this.txt_path.Location = new System.Drawing.Point(106, 66);
            this.txt_path.Name = "txt_path";
            this.txt_path.ReadOnly = true;
            this.txt_path.Size = new System.Drawing.Size(400, 34);
            this.txt_path.TabIndex = 23;
            this.txt_path.Text = "D:\\";
            // 
            // dgv_7B
            // 
            this.dgv_7B.AllowUserToAddRows = false;
            this.dgv_7B.AllowUserToDeleteRows = false;
            this.dgv_7B.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_7B.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_7B.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_7B.Location = new System.Drawing.Point(4, 5);
            this.dgv_7B.Name = "dgv_7B";
            this.dgv_7B.ReadOnly = true;
            this.dgv_7B.RowHeadersWidth = 51;
            this.dgv_7B.RowTemplate.Height = 27;
            this.dgv_7B.Size = new System.Drawing.Size(1127, 449);
            this.dgv_7B.TabIndex = 0;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tab_7B);
            this.tabControl1.Controls.Add(this.tab_INV);
            this.tabControl1.Controls.Add(this.tab_IPS);
            this.tabControl1.HotTrack = true;
            this.tabControl1.Location = new System.Drawing.Point(13, 203);
            this.tabControl1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1143, 497);
            this.tabControl1.TabIndex = 2;
            // 
            // tab_7B
            // 
            this.tab_7B.BackColor = System.Drawing.Color.Transparent;
            this.tab_7B.Controls.Add(this.dgv_7B);
            this.tab_7B.ForeColor = System.Drawing.SystemColors.ControlText;
            this.tab_7B.Location = new System.Drawing.Point(4, 34);
            this.tab_7B.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tab_7B.Name = "tab_7B";
            this.tab_7B.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tab_7B.Size = new System.Drawing.Size(1135, 459);
            this.tab_7B.TabIndex = 2;
            this.tab_7B.Text = "彙總表";
            // 
            // tab_INV
            // 
            this.tab_INV.Controls.Add(this.dgv_INV);
            this.tab_INV.Location = new System.Drawing.Point(4, 34);
            this.tab_INV.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tab_INV.Name = "tab_INV";
            this.tab_INV.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tab_INV.Size = new System.Drawing.Size(1135, 459);
            this.tab_INV.TabIndex = 3;
            this.tab_INV.Text = "庫存明細";
            this.tab_INV.UseVisualStyleBackColor = true;
            // 
            // dgv_INV
            // 
            this.dgv_INV.AllowUserToAddRows = false;
            this.dgv_INV.AllowUserToDeleteRows = false;
            this.dgv_INV.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_INV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_INV.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_INV.Location = new System.Drawing.Point(4, 5);
            this.dgv_INV.Name = "dgv_INV";
            this.dgv_INV.ReadOnly = true;
            this.dgv_INV.RowHeadersWidth = 51;
            this.dgv_INV.RowTemplate.Height = 27;
            this.dgv_INV.Size = new System.Drawing.Size(1127, 449);
            this.dgv_INV.TabIndex = 0;
            // 
            // tab_IPS
            // 
            this.tab_IPS.Controls.Add(this.dgv_IPS);
            this.tab_IPS.Location = new System.Drawing.Point(4, 34);
            this.tab_IPS.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tab_IPS.Name = "tab_IPS";
            this.tab_IPS.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.tab_IPS.Size = new System.Drawing.Size(1135, 459);
            this.tab_IPS.TabIndex = 1;
            this.tab_IPS.Text = "在途倉明細";
            this.tab_IPS.UseVisualStyleBackColor = true;
            // 
            // dgv_IPS
            // 
            this.dgv_IPS.AllowUserToAddRows = false;
            this.dgv_IPS.AllowUserToDeleteRows = false;
            this.dgv_IPS.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgv_IPS.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_IPS.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgv_IPS.Location = new System.Drawing.Point(4, 5);
            this.dgv_IPS.Name = "dgv_IPS";
            this.dgv_IPS.ReadOnly = true;
            this.dgv_IPS.RowHeadersWidth = 51;
            this.dgv_IPS.RowTemplate.Height = 27;
            this.dgv_IPS.Size = new System.Drawing.Size(1127, 449);
            this.dgv_IPS.TabIndex = 0;
            // 
            // panel2
            // 
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
            this.panel2.Location = new System.Drawing.Point(13, 12);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1138, 181);
            this.panel2.TabIndex = 3;
            // 
            // fm_Package7b
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1162, 708);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.tabControl1);
            this.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "fm_Package7b";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "關聯方期末存貨 Package7b(20200923 1510)";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.fm_Package7b_FormClosed);
            this.Load += new System.EventHandler(this.fm_Package7b_Load);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_7B)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tab_7B.ResumeLayout(false);
            this.tab_INV.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_INV)).EndInit();
            this.tab_IPS.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_IPS)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btn_up;
        private System.Windows.Forms.Button btn_down;
        private System.Windows.Forms.Button btn_fileopen;
        private System.Windows.Forms.Label lab_status;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Button btn_IPS;
        private System.Windows.Forms.Button btn_INV;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_7B;
        private System.Windows.Forms.TextBox txterr;
        private System.Windows.Forms.Button Btn_date_e;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button Btn_date_s;
        private System.Windows.Forms.TextBox txt_date_s;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btn_search;
        private System.Windows.Forms.TextBox txt_date_e;
        private System.Windows.Forms.Button btn_file;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt_path;
        private System.Windows.Forms.DataGridView dgv_7B;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tab_7B;
        private System.Windows.Forms.TabPage tab_INV;
        private System.Windows.Forms.DataGridView dgv_INV;
        private System.Windows.Forms.TabPage tab_IPS;
        private System.Windows.Forms.DataGridView dgv_IPS;
        private System.Windows.Forms.Panel panel2;
    }
}