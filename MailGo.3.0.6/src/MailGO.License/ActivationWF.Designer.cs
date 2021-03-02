namespace DataDesign.MailGO.License
{
    partial class ActivationWF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ActivationWF));
            this.pnlCover = new System.Windows.Forms.Panel();
            this.lblDesc = new System.Windows.Forms.Label();
            this.pnlMessage = new System.Windows.Forms.Panel();
            this.lblMsgMissHardware = new System.Windows.Forms.Label();
            this.lblMsgEmptyKey = new System.Windows.Forms.Label();
            this.lblMsgWarn = new System.Windows.Forms.Label();
            this.lblMsgError = new System.Windows.Forms.Label();
            this.lblMsgInvalidRequest = new System.Windows.Forms.Label();
            this.cmdGenerate = new System.Windows.Forms.Button();
            this.txtRequest = new System.Windows.Forms.TextBox();
            this.lblRequest = new System.Windows.Forms.Label();
            this.ctlSaveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.pnlCover.SuspendLayout();
            this.pnlMessage.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlCover
            // 
            this.pnlCover.Controls.Add(this.lblDesc);
            this.pnlCover.Controls.Add(this.pnlMessage);
            this.pnlCover.Controls.Add(this.cmdGenerate);
            this.pnlCover.Controls.Add(this.txtRequest);
            this.pnlCover.Controls.Add(this.lblRequest);
            resources.ApplyResources(this.pnlCover, "pnlCover");
            this.pnlCover.Name = "pnlCover";
            // 
            // lblDesc
            // 
            resources.ApplyResources(this.lblDesc, "lblDesc");
            this.lblDesc.Name = "lblDesc";
            // 
            // pnlMessage
            // 
            resources.ApplyResources(this.pnlMessage, "pnlMessage");
            this.pnlMessage.Controls.Add(this.lblMsgMissHardware);
            this.pnlMessage.Controls.Add(this.lblMsgEmptyKey);
            this.pnlMessage.Controls.Add(this.lblMsgWarn);
            this.pnlMessage.Controls.Add(this.lblMsgError);
            this.pnlMessage.Controls.Add(this.lblMsgInvalidRequest);
            this.pnlMessage.Name = "pnlMessage";
            // 
            // lblMsgMissHardware
            // 
            resources.ApplyResources(this.lblMsgMissHardware, "lblMsgMissHardware");
            this.lblMsgMissHardware.Name = "lblMsgMissHardware";
            // 
            // lblMsgEmptyKey
            // 
            resources.ApplyResources(this.lblMsgEmptyKey, "lblMsgEmptyKey");
            this.lblMsgEmptyKey.Name = "lblMsgEmptyKey";
            // 
            // lblMsgWarn
            // 
            resources.ApplyResources(this.lblMsgWarn, "lblMsgWarn");
            this.lblMsgWarn.Name = "lblMsgWarn";
            // 
            // lblMsgError
            // 
            resources.ApplyResources(this.lblMsgError, "lblMsgError");
            this.lblMsgError.Name = "lblMsgError";
            // 
            // lblMsgInvalidRequest
            // 
            resources.ApplyResources(this.lblMsgInvalidRequest, "lblMsgInvalidRequest");
            this.lblMsgInvalidRequest.Name = "lblMsgInvalidRequest";
            // 
            // cmdGenerate
            // 
            resources.ApplyResources(this.cmdGenerate, "cmdGenerate");
            this.cmdGenerate.Name = "cmdGenerate";
            this.cmdGenerate.UseVisualStyleBackColor = true;
            this.cmdGenerate.Click += new System.EventHandler(this.cmdGenerate_Click);
            // 
            // txtRequest
            // 
            resources.ApplyResources(this.txtRequest, "txtRequest");
            this.txtRequest.Name = "txtRequest";
            // 
            // lblRequest
            // 
            resources.ApplyResources(this.lblRequest, "lblRequest");
            this.lblRequest.Name = "lblRequest";
            // 
            // ctlSaveFileDialog
            // 
            resources.ApplyResources(this.ctlSaveFileDialog, "ctlSaveFileDialog");
            // 
            // ActivationWF
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pnlCover);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ActivationWF";
            this.pnlCover.ResumeLayout(false);
            this.pnlCover.PerformLayout();
            this.pnlMessage.ResumeLayout(false);
            this.pnlMessage.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Label lblRequest;
        private System.Windows.Forms.TextBox txtRequest;
        private System.Windows.Forms.Button cmdGenerate;
        private System.Windows.Forms.SaveFileDialog ctlSaveFileDialog;
        private System.Windows.Forms.Panel pnlMessage;
        private System.Windows.Forms.Label lblMsgError;
        private System.Windows.Forms.Label lblMsgInvalidRequest;
        private System.Windows.Forms.Label lblDesc;
        private System.Windows.Forms.Label lblMsgWarn;
        private System.Windows.Forms.Label lblMsgEmptyKey;
        private System.Windows.Forms.Label lblMsgMissHardware;
    }
}