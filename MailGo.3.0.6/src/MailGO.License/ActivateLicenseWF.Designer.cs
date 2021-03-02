namespace DataDesign.MailGO.License
{
    partial class ActivateLicenseWF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ActivateLicenseWF));
            this.pnlCover = new System.Windows.Forms.Panel();
            this.pnlMessage = new System.Windows.Forms.Panel();
            this.lblMsgError = new System.Windows.Forms.Label();
            this.lblMsgInvalidActivation = new System.Windows.Forms.Label();
            this.cmdOpen = new System.Windows.Forms.Button();
            this.txtActivationID = new System.Windows.Forms.TextBox();
            this.lblActivationID = new System.Windows.Forms.Label();
            this.cmdActivate = new System.Windows.Forms.Button();
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
            this.pnlCover.Controls.Add(this.txtActivationID);
            this.pnlCover.Controls.Add(this.lblActivationID);
            this.pnlCover.Controls.Add(this.cmdActivate);
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
            this.pnlMessage.Controls.Add(this.lblMsgInvalidActivation);
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
            // lblMsgInvalidActivation
            // 
            this.lblMsgInvalidActivation.AccessibleDescription = null;
            this.lblMsgInvalidActivation.AccessibleName = null;
            resources.ApplyResources(this.lblMsgInvalidActivation, "lblMsgInvalidActivation");
            this.lblMsgInvalidActivation.Font = null;
            this.lblMsgInvalidActivation.Name = "lblMsgInvalidActivation";
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
            // txtActivationID
            // 
            this.txtActivationID.AccessibleDescription = null;
            this.txtActivationID.AccessibleName = null;
            resources.ApplyResources(this.txtActivationID, "txtActivationID");
            this.txtActivationID.BackColor = System.Drawing.SystemColors.Window;
            this.txtActivationID.BackgroundImage = null;
            this.txtActivationID.Font = null;
            this.txtActivationID.Name = "txtActivationID";
            // 
            // lblActivationID
            // 
            this.lblActivationID.AccessibleDescription = null;
            this.lblActivationID.AccessibleName = null;
            resources.ApplyResources(this.lblActivationID, "lblActivationID");
            this.lblActivationID.Font = null;
            this.lblActivationID.Name = "lblActivationID";
            // 
            // cmdActivate
            // 
            this.cmdActivate.AccessibleDescription = null;
            this.cmdActivate.AccessibleName = null;
            resources.ApplyResources(this.cmdActivate, "cmdActivate");
            this.cmdActivate.BackgroundImage = null;
            this.cmdActivate.Font = null;
            this.cmdActivate.Name = "cmdActivate";
            this.cmdActivate.UseVisualStyleBackColor = true;
            this.cmdActivate.Click += new System.EventHandler(this.cmdActivate_Click);
            // 
            // ctlOpenFileDialog
            // 
            resources.ApplyResources(this.ctlOpenFileDialog, "ctlOpenFileDialog");
            // 
            // ActivateLicenseWF
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
            this.Name = "ActivateLicenseWF";
            this.pnlCover.ResumeLayout(false);
            this.pnlCover.PerformLayout();
            this.pnlMessage.ResumeLayout(false);
            this.pnlMessage.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Button cmdOpen;
        private System.Windows.Forms.TextBox txtActivationID;
        private System.Windows.Forms.Label lblActivationID;
        private System.Windows.Forms.Button cmdActivate;
        private System.Windows.Forms.OpenFileDialog ctlOpenFileDialog;
        private System.Windows.Forms.Panel pnlMessage;
        private System.Windows.Forms.Label lblMsgError;
        private System.Windows.Forms.Label lblMsgInvalidActivation;
    }
}