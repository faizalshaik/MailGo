namespace DataDesign.MailGO.Activation
{
    partial class NotMatchWF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NotMatchWF));
            this.pnlCover = new System.Windows.Forms.Panel();
            this.lblProperty = new System.Windows.Forms.Label();
            this.lblFileName = new System.Windows.Forms.Label();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.cmdSend = new System.Windows.Forms.Button();
            this.lblFileProperty = new System.Windows.Forms.Label();
            this.lblAddress = new System.Windows.Forms.Label();
            this.lblMessage = new System.Windows.Forms.Label();
            this.pnlCover.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlCover
            // 
            this.pnlCover.Controls.Add(this.lblProperty);
            this.pnlCover.Controls.Add(this.lblFileName);
            this.pnlCover.Controls.Add(this.cmdCancel);
            this.pnlCover.Controls.Add(this.cmdSend);
            this.pnlCover.Controls.Add(this.lblFileProperty);
            this.pnlCover.Controls.Add(this.lblAddress);
            this.pnlCover.Controls.Add(this.lblMessage);
            resources.ApplyResources(this.pnlCover, "pnlCover");
            this.pnlCover.Name = "pnlCover";
            // 
            // lblProperty
            // 
            resources.ApplyResources(this.lblProperty, "lblProperty");
            this.lblProperty.Name = "lblProperty";
            // 
            // lblFileName
            // 
            resources.ApplyResources(this.lblFileName, "lblFileName");
            this.lblFileName.Name = "lblFileName";
            // 
            // cmdCancel
            // 
            resources.ApplyResources(this.cmdCancel, "cmdCancel");
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.UseVisualStyleBackColor = true;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // cmdSend
            // 
            resources.ApplyResources(this.cmdSend, "cmdSend");
            this.cmdSend.Name = "cmdSend";
            this.cmdSend.UseVisualStyleBackColor = true;
            this.cmdSend.Click += new System.EventHandler(this.cmdSend_Click);
            // 
            // lblFileProperty
            // 
            resources.ApplyResources(this.lblFileProperty, "lblFileProperty");
            this.lblFileProperty.Name = "lblFileProperty";
            // 
            // lblAddress
            // 
            resources.ApplyResources(this.lblAddress, "lblAddress");
            this.lblAddress.Name = "lblAddress";
            // 
            // lblMessage
            // 
            resources.ApplyResources(this.lblMessage, "lblMessage");
            this.lblMessage.Name = "lblMessage";
            // 
            // NotMatchWF
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pnlCover);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "NotMatchWF";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.NotMatchWF_FormClosing);
            this.Load += new System.EventHandler(this.NotMatchWF_Load);
            this.pnlCover.ResumeLayout(false);
            this.pnlCover.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Label lblMessage;
        private System.Windows.Forms.Label lblAddress;
        private System.Windows.Forms.Label lblFileProperty;
        private System.Windows.Forms.Button cmdSend;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.Label lblFileName;
        private System.Windows.Forms.Label lblProperty;
    }
}