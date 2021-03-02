namespace DataDesign.MailGO.License
{
    partial class InstallLicenseWF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(InstallLicenseWF));
            this.pnlCover = new System.Windows.Forms.Panel();
            this.pnlMessage = new System.Windows.Forms.Panel();
            this.lblMsgError = new System.Windows.Forms.Label();
            this.lblMsgInvalidLicense = new System.Windows.Forms.Label();
            this.cmdOpen = new System.Windows.Forms.Button();
            this.txtLicenseID = new System.Windows.Forms.TextBox();
            this.lblLicenseID = new System.Windows.Forms.Label();
            this.cmdInstall = new System.Windows.Forms.Button();
            this.ctlOpenFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.pnlCover.SuspendLayout();
            this.pnlMessage.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlCover
            // 
            this.pnlCover.AccessibleDescription = null;
            this.pnlCover.AccessibleName = null;
            resources.ApplyResources(this.pnlCover, "pnlCover");
            this.pnlCover.BackgroundImage = null;
            this.pnlCover.Controls.Add(this.pnlMessage);
            this.pnlCover.Controls.Add(this.cmdOpen);
            this.pnlCover.Controls.Add(this.txtLicenseID);
            this.pnlCover.Controls.Add(this.lblLicenseID);
            this.pnlCover.Controls.Add(this.cmdInstall);
            this.pnlCover.Font = null;
            this.pnlCover.Name = "pnlCover";
            // 
            // pnlMessage
            // 
            this.pnlMessage.AccessibleDescription = null;
            this.pnlMessage.AccessibleName = null;
            resources.ApplyResources(this.pnlMessage, "pnlMessage");
            this.pnlMessage.BackgroundImage = null;
            this.pnlMessage.Controls.Add(this.lblMsgError);
            this.pnlMessage.Controls.Add(this.lblMsgInvalidLicense);
            this.pnlMessage.Font = null;
            this.pnlMessage.Name = "pnlMessage";
            // 
            // lblMsgError
            // 
            this.lblMsgError.AccessibleDescription = null;
            this.lblMsgError.AccessibleName = null;
            resources.ApplyResources(this.lblMsgError, "lblMsgError");
            this.lblMsgError.Font = null;
            this.lblMsgError.Name = "lblMsgError";
            // 
            // lblMsgInvalidLicense
            // 
            this.lblMsgInvalidLicense.AccessibleDescription = null;
            this.lblMsgInvalidLicense.AccessibleName = null;
            resources.ApplyResources(this.lblMsgInvalidLicense, "lblMsgInvalidLicense");
            this.lblMsgInvalidLicense.Font = null;
            this.lblMsgInvalidLicense.Name = "lblMsgInvalidLicense";
            // 
            // cmdOpen
            // 
            this.cmdOpen.AccessibleDescription = null;
            this.cmdOpen.AccessibleName = null;
            resources.ApplyResources(this.cmdOpen, "cmdOpen");
            this.cmdOpen.BackgroundImage = null;
            this.cmdOpen.Font = null;
            this.cmdOpen.Name = "cmdOpen";
            this.cmdOpen.UseVisualStyleBackColor = true;
            this.cmdOpen.Click += new System.EventHandler(this.cmdOpen_Click);
            // 
            // txtLicenseID
            // 
            this.txtLicenseID.AccessibleDescription = null;
            this.txtLicenseID.AccessibleName = null;
            resources.ApplyResources(this.txtLicenseID, "txtLicenseID");
            this.txtLicenseID.BackColor = System.Drawing.SystemColors.Window;
            this.txtLicenseID.BackgroundImage = null;
            this.txtLicenseID.Font = null;
            this.txtLicenseID.Name = "txtLicenseID";
            // 
            // lblLicenseID
            // 
            this.lblLicenseID.AccessibleDescription = null;
            this.lblLicenseID.AccessibleName = null;
            resources.ApplyResources(this.lblLicenseID, "lblLicenseID");
            this.lblLicenseID.Font = null;
            this.lblLicenseID.Name = "lblLicenseID";
            // 
            // cmdInstall
            // 
            this.cmdInstall.AccessibleDescription = null;
            this.cmdInstall.AccessibleName = null;
            resources.ApplyResources(this.cmdInstall, "cmdInstall");
            this.cmdInstall.BackgroundImage = null;
            this.cmdInstall.Font = null;
            this.cmdInstall.Name = "cmdInstall";
            this.cmdInstall.UseVisualStyleBackColor = true;
            this.cmdInstall.Click += new System.EventHandler(this.cmdInstall_Click);
            // 
            // ctlOpenFileDialog
            // 
            resources.ApplyResources(this.ctlOpenFileDialog, "ctlOpenFileDialog");
            // 
            // InstallLicenseWF
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
            this.Name = "InstallLicenseWF";
            this.pnlCover.ResumeLayout(false);
            this.pnlCover.PerformLayout();
            this.pnlMessage.ResumeLayout(false);
            this.pnlMessage.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Button cmdOpen;
        private System.Windows.Forms.TextBox txtLicenseID;
        private System.Windows.Forms.Label lblLicenseID;
        private System.Windows.Forms.Button cmdInstall;
        private System.Windows.Forms.OpenFileDialog ctlOpenFileDialog;
        private System.Windows.Forms.Panel pnlMessage;
        private System.Windows.Forms.Label lblMsgError;
        private System.Windows.Forms.Label lblMsgInvalidLicense;
    }
}