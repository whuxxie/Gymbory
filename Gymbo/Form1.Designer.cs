namespace Gymbo
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnopen = new DevExpress.XtraEditors.SimpleButton();
            this.btngenerate = new DevExpress.XtraEditors.SimpleButton();
            this.meenamelist = new DevExpress.XtraEditors.MemoEdit();
            this.btnsavename = new DevExpress.XtraEditors.SimpleButton();
            this.btnreflash = new DevExpress.XtraEditors.SimpleButton();
            this.spreadsheetControl1 = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
            this.xtc = new DevExpress.XtraTab.XtraTabControl();
            this.xtpschedule = new DevExpress.XtraTab.XtraTabPage();
            this.xtpstatistics = new DevExpress.XtraTab.XtraTabPage();
            this.spreadsheetControl2 = new DevExpress.XtraSpreadsheet.SpreadsheetControl();
            this.btncheck = new DevExpress.XtraEditors.SimpleButton();
            this.btnimportstas = new DevExpress.XtraEditors.SimpleButton();
            this.btnexport = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.meenamelist.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.xtc)).BeginInit();
            this.xtc.SuspendLayout();
            this.xtpschedule.SuspendLayout();
            this.xtpstatistics.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnopen
            // 
            this.btnopen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnopen.Location = new System.Drawing.Point(358, 556);
            this.btnopen.LookAndFeel.SkinName = "VS2010";
            this.btnopen.LookAndFeel.UseDefaultLookAndFeel = false;
            this.btnopen.Name = "btnopen";
            this.btnopen.Size = new System.Drawing.Size(99, 23);
            this.btnopen.TabIndex = 0;
            this.btnopen.Text = "打开Schedule";
            this.btnopen.Click += new System.EventHandler(this.btnopen_Click);
            // 
            // btngenerate
            // 
            this.btngenerate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btngenerate.Location = new System.Drawing.Point(778, 556);
            this.btngenerate.LookAndFeel.SkinName = "VS2010";
            this.btngenerate.LookAndFeel.UseDefaultLookAndFeel = false;
            this.btngenerate.Name = "btngenerate";
            this.btngenerate.Size = new System.Drawing.Size(99, 23);
            this.btngenerate.TabIndex = 2;
            this.btngenerate.Text = "生成统计表";
            this.btngenerate.Click += new System.EventHandler(this.btngenerate_Click);
            // 
            // meenamelist
            // 
            this.meenamelist.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.meenamelist.EditValue = "";
            this.meenamelist.Location = new System.Drawing.Point(12, 12);
            this.meenamelist.Name = "meenamelist";
            this.meenamelist.Properties.Appearance.Font = new System.Drawing.Font("Adobe Caslon Pro", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.meenamelist.Properties.Appearance.Options.UseFont = true;
            this.meenamelist.Size = new System.Drawing.Size(137, 528);
            this.meenamelist.TabIndex = 3;
            this.meenamelist.UseOptimizedRendering = true;
            // 
            // btnsavename
            // 
            this.btnsavename.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnsavename.Location = new System.Drawing.Point(74, 556);
            this.btnsavename.Name = "btnsavename";
            this.btnsavename.Size = new System.Drawing.Size(75, 23);
            this.btnsavename.TabIndex = 4;
            this.btnsavename.Text = "保存名单";
            this.btnsavename.Click += new System.EventHandler(this.btnsavename_Click);
            // 
            // btnreflash
            // 
            this.btnreflash.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnreflash.Location = new System.Drawing.Point(12, 556);
            this.btnreflash.Name = "btnreflash";
            this.btnreflash.Size = new System.Drawing.Size(56, 23);
            this.btnreflash.TabIndex = 5;
            this.btnreflash.Text = "刷新";
            this.btnreflash.Click += new System.EventHandler(this.btnreflash_Click);
            // 
            // spreadsheetControl1
            // 
            this.spreadsheetControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.spreadsheetControl1.Location = new System.Drawing.Point(0, 0);
            this.spreadsheetControl1.Name = "spreadsheetControl1";
            this.spreadsheetControl1.Size = new System.Drawing.Size(937, 499);
            this.spreadsheetControl1.TabIndex = 6;
            this.spreadsheetControl1.Text = "spreadsheetControl1";
            // 
            // xtc
            // 
            this.xtc.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.xtc.Location = new System.Drawing.Point(155, 12);
            this.xtc.Name = "xtc";
            this.xtc.SelectedTabPage = this.xtpschedule;
            this.xtc.Size = new System.Drawing.Size(943, 528);
            this.xtc.TabIndex = 7;
            this.xtc.TabPages.AddRange(new DevExpress.XtraTab.XtraTabPage[] {
            this.xtpschedule,
            this.xtpstatistics});
            // 
            // xtpschedule
            // 
            this.xtpschedule.Controls.Add(this.spreadsheetControl1);
            this.xtpschedule.Name = "xtpschedule";
            this.xtpschedule.Size = new System.Drawing.Size(937, 499);
            this.xtpschedule.Text = "课程表";
            // 
            // xtpstatistics
            // 
            this.xtpstatistics.Controls.Add(this.spreadsheetControl2);
            this.xtpstatistics.Name = "xtpstatistics";
            this.xtpstatistics.Size = new System.Drawing.Size(1051, 469);
            this.xtpstatistics.Text = "课时统计表";
            // 
            // spreadsheetControl2
            // 
            this.spreadsheetControl2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.spreadsheetControl2.Location = new System.Drawing.Point(0, 0);
            this.spreadsheetControl2.Name = "spreadsheetControl2";
            this.spreadsheetControl2.Size = new System.Drawing.Size(1051, 469);
            this.spreadsheetControl2.TabIndex = 0;
            this.spreadsheetControl2.Text = "spreadsheetControl2";
            // 
            // btncheck
            // 
            this.btncheck.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btncheck.Location = new System.Drawing.Point(498, 556);
            this.btncheck.LookAndFeel.SkinName = "VS2010";
            this.btncheck.LookAndFeel.UseDefaultLookAndFeel = false;
            this.btncheck.Name = "btncheck";
            this.btncheck.Size = new System.Drawing.Size(99, 23);
            this.btncheck.TabIndex = 8;
            this.btncheck.Text = "检查格式";
            this.btncheck.Click += new System.EventHandler(this.btncheck_Click);
            // 
            // btnimportstas
            // 
            this.btnimportstas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnimportstas.Location = new System.Drawing.Point(638, 556);
            this.btnimportstas.LookAndFeel.SkinName = "VS2010";
            this.btnimportstas.LookAndFeel.UseDefaultLookAndFeel = false;
            this.btnimportstas.Name = "btnimportstas";
            this.btnimportstas.Size = new System.Drawing.Size(99, 23);
            this.btnimportstas.TabIndex = 9;
            this.btnimportstas.Text = "导入课时统计表";
            this.btnimportstas.Click += new System.EventHandler(this.btnimportstas_Click);
            // 
            // btnexport
            // 
            this.btnexport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnexport.Location = new System.Drawing.Point(918, 556);
            this.btnexport.LookAndFeel.SkinName = "VS2010";
            this.btnexport.LookAndFeel.UseDefaultLookAndFeel = false;
            this.btnexport.Name = "btnexport";
            this.btnexport.Size = new System.Drawing.Size(99, 23);
            this.btnexport.TabIndex = 10;
            this.btnexport.Text = "导出统计表";
            this.btnexport.Click += new System.EventHandler(this.btnexport_Click);
            // 
            // Form1
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1099, 584);
            this.Controls.Add(this.btnexport);
            this.Controls.Add(this.btnimportstas);
            this.Controls.Add(this.btncheck);
            this.Controls.Add(this.xtc);
            this.Controls.Add(this.btnreflash);
            this.Controls.Add(this.btnsavename);
            this.Controls.Add(this.btngenerate);
            this.Controls.Add(this.btnopen);
            this.Controls.Add(this.meenamelist);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.LookAndFeel.SkinName = "Office 2010 Blue";
            this.LookAndFeel.UseDefaultLookAndFeel = false;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "Form1";
            this.Text = "课时统计";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.meenamelist.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.xtc)).EndInit();
            this.xtc.ResumeLayout(false);
            this.xtpschedule.ResumeLayout(false);
            this.xtpstatistics.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.SimpleButton btnopen;
        private DevExpress.XtraEditors.SimpleButton btngenerate;
        private DevExpress.XtraEditors.MemoEdit meenamelist;
        private DevExpress.XtraEditors.SimpleButton btnsavename;
        private DevExpress.XtraEditors.SimpleButton btnreflash;
        private DevExpress.XtraSpreadsheet.SpreadsheetControl spreadsheetControl1;
        private DevExpress.XtraTab.XtraTabControl xtc;
        private DevExpress.XtraTab.XtraTabPage xtpschedule;
        private DevExpress.XtraTab.XtraTabPage xtpstatistics;
        private DevExpress.XtraSpreadsheet.SpreadsheetControl spreadsheetControl2;
        private DevExpress.XtraEditors.SimpleButton btncheck;
        private DevExpress.XtraEditors.SimpleButton btnimportstas;
        private DevExpress.XtraEditors.SimpleButton btnexport;


    }
}

