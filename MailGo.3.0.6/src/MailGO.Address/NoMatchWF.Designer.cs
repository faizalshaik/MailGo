namespace DataDesign.MailGO.Address
{
    partial class NoMatchWF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NoMatchWF));
            this.pnlCover = new System.Windows.Forms.Panel();
            this.cmdIgnore = new System.Windows.Forms.Button();
            this.cmdAddList = new System.Windows.Forms.Button();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.lblCompany = new System.Windows.Forms.Label();
            this.lblEmail = new System.Windows.Forms.Label();
            this.lblMessage = new System.Windows.Forms.Label();
            this.pnlCover.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlCover
            // 
            this.pnlCover.AccessibleDescription = null;
            this.pnlCover.AccessibleName = null;
            resources.ApplyResources(this.pnlCover, "pnlCover");
            this.pnlCover.BackgroundImage = null;
            this.pnlCover.Controls.Add(this.cmdIgnore);
            this.pnlCover.Controls.Add(this.cmdAddList);
            this.pnlCover.Controls.Add(this.cmdCancel);
            this.pnlCover.Controls.Add(this.lblCompany);
            this.pnlCover.Controls.Add(this.lblEmail);
            this.pnlCover.Controls.Add(this.lblMessage);
            this.pnlCover.Font = null;
            this.pnlCover.Name = "pnlCover";
            // 
            // cmdIgnore
            // 
            this.cmdIgnore.AccessibleDescription = null;
            this.cmdIgnore.AccessibleName = null;
            resources.ApplyResources(this.cmdIgnore, "cmdIgnore");
            this.cmdIgnore.BackgroundImage = null;
            this.cmdIgnore.Font = null;
            this.cmdIgnore.Name = "cmdIgnore";
            this.cmdIgnore.UseVisualStyleBackColor = true;
            this.cmdIgnore.Click += new System.EventHandler(this.cmdIgnore_Click);
            // 
            // cmdAddList
            // 
            this.cmdAddList.AccessibleDescription = null;
            this.cmdAddList.AccessibleName = null;
            resources.ApplyResources(this.cmdAddList, "cmdAddList");
            this.cmdAddList.BackgroundImage = null;
            this.cmdAddList.Font = null;
            this.cmdAddList.Name = "cmdAddList";
            this.cmdAddList.UseVisualStyleBackColor = true;
            this.cmdAddList.Click += new System.EventHandler(this.cmdAddList_Click);
            // 
            // cmdCancel
            // 
            this.cmdCancel.AccessibleDescription = null;
            this.cmdCancel.AccessibleName = null;
            resources.ApplyResources(this.cmdCancel, "cmdCancel");
            this.cmdCancel.BackgroundImage = null;
            this.cmdCancel.Font = null;
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.UseVisualStyleBackColor = true;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // lblCompany
            // 
            this.lblCompany.AccessibleDescription = null;
            this.lblCompany.AccessibleName = null;
            resources.ApplyResources(this.lblCompany, "lblCompany");
            this.lblCompany.Font = null;
            this.lblCompany.Name = "lblCompany";
            // 
            // lblEmail
            // 
            this.lblEmail.AccessibleDescription = null;
            this.lblEmail.AccessibleName = null;
            resources.ApplyResources(this.lblEmail, "lblEmail");
            this.lblEmail.Font = null;
            this.lblEmail.Name = "lblEmail";
            // 
            // lblMessage
            // 
            this.lblMessage.AccessibleDescription = null;
            this.lblMessage.AccessibleName = null;
            resources.ApplyResources(this.lblMessage, "lblMessage");
            this.lblMessage.Font = null;
            this.lblMessage.Name = "lblMessage";
            // 
            // NoMatchWF
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
            this.Name = "NoMatchWF";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.NoMatchWF_FormClosing);
            this.Load += new System.EventHandler(this.NoMatchWF_Load);
            this.pnlCover.ResumeLayout(false);
            this.pnlCover.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Label lblMessage;
        private System.Windows.Forms.Label lblEmail;
        private System.Windows.Forms.Label lblCompany;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.Button cmdAddList;
        private System.Windows.Forms.Button cmdIgnore;
    }
}