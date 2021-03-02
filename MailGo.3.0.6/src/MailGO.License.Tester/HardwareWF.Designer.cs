namespace DataDesign.MailGO.License.Tester
{
    partial class HardwareWF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(HardwareWF));
            this.pnlCover = new System.Windows.Forms.Panel();
            this.txtCPU = new System.Windows.Forms.TextBox();
            this.lblCPU = new System.Windows.Forms.Label();
            this.txtMAC = new System.Windows.Forms.TextBox();
            this.lblMAC = new System.Windows.Forms.Label();
            this.pnlCover.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlCover
            // 
            this.pnlCover.AccessibleDescription = null;
            this.pnlCover.AccessibleName = null;
            resources.ApplyResources(this.pnlCover, "pnlCover");
            this.pnlCover.BackgroundImage = null;
            this.pnlCover.Controls.Add(this.txtCPU);
            this.pnlCover.Controls.Add(this.lblCPU);
            this.pnlCover.Controls.Add(this.txtMAC);
            this.pnlCover.Controls.Add(this.lblMAC);
            this.pnlCover.Font = null;
            this.pnlCover.Name = "pnlCover";
            // 
            // txtCPU
            // 
            this.txtCPU.AccessibleDescription = null;
            this.txtCPU.AccessibleName = null;
            resources.ApplyResources(this.txtCPU, "txtCPU");
            this.txtCPU.BackgroundImage = null;
            this.txtCPU.Font = null;
            this.txtCPU.Name = "txtCPU";
            this.txtCPU.ReadOnly = true;
            // 
            // lblCPU
            // 
            this.lblCPU.AccessibleDescription = null;
            this.lblCPU.AccessibleName = null;
            resources.ApplyResources(this.lblCPU, "lblCPU");
            this.lblCPU.Font = null;
            this.lblCPU.Name = "lblCPU";
            // 
            // txtMAC
            // 
            this.txtMAC.AccessibleDescription = null;
            this.txtMAC.AccessibleName = null;
            resources.ApplyResources(this.txtMAC, "txtMAC");
            this.txtMAC.BackgroundImage = null;
            this.txtMAC.Font = null;
            this.txtMAC.Name = "txtMAC";
            this.txtMAC.ReadOnly = true;
            // 
            // lblMAC
            // 
            this.lblMAC.AccessibleDescription = null;
            this.lblMAC.AccessibleName = null;
            resources.ApplyResources(this.lblMAC, "lblMAC");
            this.lblMAC.Font = null;
            this.lblMAC.Name = "lblMAC";
            // 
            // HardwareWF
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
            this.Name = "HardwareWF";
            this.Load += new System.EventHandler(this.HardwareWF_Load);
            this.pnlCover.ResumeLayout(false);
            this.pnlCover.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Label lblMAC;
        private System.Windows.Forms.TextBox txtMAC;
        private System.Windows.Forms.TextBox txtCPU;
        private System.Windows.Forms.Label lblCPU;
    }
}