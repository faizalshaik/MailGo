namespace DataDesign.MailGO.Activation
{
    partial class AskActivationWF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AskActivationWF));
            this.pnlCover = new System.Windows.Forms.Panel();
            this.cmdNo = new System.Windows.Forms.Button();
            this.cmdYes = new System.Windows.Forms.Button();
            this.lblMessage = new System.Windows.Forms.Label();
            this.pnlCover.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlCover
            // 
            this.pnlCover.Controls.Add(this.cmdNo);
            this.pnlCover.Controls.Add(this.cmdYes);
            this.pnlCover.Controls.Add(this.lblMessage);
            resources.ApplyResources(this.pnlCover, "pnlCover");
            this.pnlCover.Name = "pnlCover";
            // 
            // cmdNo
            // 
            resources.ApplyResources(this.cmdNo, "cmdNo");
            this.cmdNo.Name = "cmdNo";
            this.cmdNo.UseVisualStyleBackColor = true;
            this.cmdNo.Click += new System.EventHandler(this.cmdNo_Click);
            // 
            // cmdYes
            // 
            resources.ApplyResources(this.cmdYes, "cmdYes");
            this.cmdYes.Name = "cmdYes";
            this.cmdYes.UseVisualStyleBackColor = true;
            this.cmdYes.Click += new System.EventHandler(this.cmdYes_Click);
            // 
            // lblMessage
            // 
            resources.ApplyResources(this.lblMessage, "lblMessage");
            this.lblMessage.Name = "lblMessage";
            // 
            // AskActivationWF
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pnlCover);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AskActivationWF";
            this.Load += new System.EventHandler(this.AskActivationWF_Load);
            this.pnlCover.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Label lblMessage;
        private System.Windows.Forms.Button cmdNo;
        private System.Windows.Forms.Button cmdYes;
    }
}