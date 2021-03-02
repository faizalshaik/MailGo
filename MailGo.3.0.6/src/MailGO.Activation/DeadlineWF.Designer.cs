namespace DataDesign.MailGO.Activation
{
    partial class DeadlineWF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DeadlineWF));
            this.pnlCover = new System.Windows.Forms.Panel();
            this.cmdOK = new System.Windows.Forms.Button();
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
            this.pnlCover.Controls.Add(this.cmdOK);
            this.pnlCover.Controls.Add(this.lblMessage);
            this.pnlCover.Font = null;
            this.pnlCover.Name = "pnlCover";
            // 
            // cmdOK
            // 
            this.cmdOK.AccessibleDescription = null;
            this.cmdOK.AccessibleName = null;
            resources.ApplyResources(this.cmdOK, "cmdOK");
            this.cmdOK.BackgroundImage = null;
            this.cmdOK.Font = null;
            this.cmdOK.Name = "cmdOK";
            this.cmdOK.UseVisualStyleBackColor = true;
            this.cmdOK.Click += new System.EventHandler(this.cmdOK_Click);
            // 
            // lblMessage
            // 
            this.lblMessage.AccessibleDescription = null;
            this.lblMessage.AccessibleName = null;
            resources.ApplyResources(this.lblMessage, "lblMessage");
            this.lblMessage.Font = null;
            this.lblMessage.Name = "lblMessage";
            // 
            // DeadlineWF
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
            this.Name = "DeadlineWF";
            this.Load += new System.EventHandler(this.DeadlineWF_Load);
            this.pnlCover.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Label lblMessage;
        private System.Windows.Forms.Button cmdOK;
    }
}