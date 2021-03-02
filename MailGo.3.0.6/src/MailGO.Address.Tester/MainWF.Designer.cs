namespace DataDesign.MailGO.Address.Tester
{
    partial class MainWF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWF));
            this.pnlCover = new System.Windows.Forms.Panel();
            this.cboLanguage = new System.Windows.Forms.ComboBox();
            this.lblLanguage = new System.Windows.Forms.Label();
            this.cmdCheckEmail = new System.Windows.Forms.Button();
            this.cmdCreateList = new System.Windows.Forms.Button();
            this.pnlCover.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlCover
            // 
            this.pnlCover.AccessibleDescription = null;
            this.pnlCover.AccessibleName = null;
            resources.ApplyResources(this.pnlCover, "pnlCover");
            this.pnlCover.BackgroundImage = null;
            this.pnlCover.Controls.Add(this.cboLanguage);
            this.pnlCover.Controls.Add(this.lblLanguage);
            this.pnlCover.Controls.Add(this.cmdCheckEmail);
            this.pnlCover.Controls.Add(this.cmdCreateList);
            this.pnlCover.Font = null;
            this.pnlCover.Name = "pnlCover";
            // 
            // cboLanguage
            // 
            this.cboLanguage.AccessibleDescription = null;
            this.cboLanguage.AccessibleName = null;
            resources.ApplyResources(this.cboLanguage, "cboLanguage");
            this.cboLanguage.BackgroundImage = null;
            this.cboLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboLanguage.Font = null;
            this.cboLanguage.FormattingEnabled = true;
            this.cboLanguage.Items.AddRange(new object[] {
            resources.GetString("cboLanguage.Items"),
            resources.GetString("cboLanguage.Items1")});
            this.cboLanguage.Name = "cboLanguage";
            this.cboLanguage.SelectedIndexChanged += new System.EventHandler(this.cboLanguage_SelectedIndexChanged);
            // 
            // lblLanguage
            // 
            this.lblLanguage.AccessibleDescription = null;
            this.lblLanguage.AccessibleName = null;
            resources.ApplyResources(this.lblLanguage, "lblLanguage");
            this.lblLanguage.Font = null;
            this.lblLanguage.Name = "lblLanguage";
            // 
            // cmdCheckEmail
            // 
            this.cmdCheckEmail.AccessibleDescription = null;
            this.cmdCheckEmail.AccessibleName = null;
            resources.ApplyResources(this.cmdCheckEmail, "cmdCheckEmail");
            this.cmdCheckEmail.BackgroundImage = null;
            this.cmdCheckEmail.Font = null;
            this.cmdCheckEmail.Name = "cmdCheckEmail";
            this.cmdCheckEmail.UseVisualStyleBackColor = true;
            this.cmdCheckEmail.Click += new System.EventHandler(this.cmdCheckEmail_Click);
            // 
            // cmdCreateList
            // 
            this.cmdCreateList.AccessibleDescription = null;
            this.cmdCreateList.AccessibleName = null;
            resources.ApplyResources(this.cmdCreateList, "cmdCreateList");
            this.cmdCreateList.BackgroundImage = null;
            this.cmdCreateList.Font = null;
            this.cmdCreateList.Name = "cmdCreateList";
            this.cmdCreateList.UseVisualStyleBackColor = true;
            this.cmdCreateList.Click += new System.EventHandler(this.cmdCreateList_Click);
            // 
            // MainWF
            // 
            this.AccessibleDescription = null;
            this.AccessibleName = null;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = null;
            this.Controls.Add(this.pnlCover);
            this.Font = null;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainWF";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainWF_FormClosed);
            this.Load += new System.EventHandler(this.MainWF_Load);
            this.pnlCover.ResumeLayout(false);
            this.pnlCover.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Button cmdCreateList;
        private System.Windows.Forms.Button cmdCheckEmail;
        private System.Windows.Forms.Label lblLanguage;
        private System.Windows.Forms.ComboBox cboLanguage;
    }
}