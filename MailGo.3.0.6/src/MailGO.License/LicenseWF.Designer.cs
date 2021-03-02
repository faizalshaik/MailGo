namespace DataDesign.MailGO.License
{
    partial class LicenseWF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LicenseWF));
            this.pnlCover = new System.Windows.Forms.Panel();
            this.cmdGenerate = new System.Windows.Forms.Button();
            this.lblTimeUnit = new System.Windows.Forms.Label();
            this.nudActivateTime = new System.Windows.Forms.NumericUpDown();
            this.lblActivateTime = new System.Windows.Forms.Label();
            this.ctlSaveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.pnlCover.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudActivateTime)).BeginInit();
            this.SuspendLayout();
            // 
            // pnlCover
            // 
            this.pnlCover.AccessibleDescription = null;
            this.pnlCover.AccessibleName = null;
            resources.ApplyResources(this.pnlCover, "pnlCover");
            this.pnlCover.BackgroundImage = null;
            this.pnlCover.Controls.Add(this.cmdGenerate);
            this.pnlCover.Controls.Add(this.lblTimeUnit);
            this.pnlCover.Controls.Add(this.nudActivateTime);
            this.pnlCover.Controls.Add(this.lblActivateTime);
            this.pnlCover.Font = null;
            this.pnlCover.Name = "pnlCover";
            // 
            // cmdGenerate
            // 
            this.cmdGenerate.AccessibleDescription = null;
            this.cmdGenerate.AccessibleName = null;
            resources.ApplyResources(this.cmdGenerate, "cmdGenerate");
            this.cmdGenerate.BackgroundImage = null;
            this.cmdGenerate.Font = null;
            this.cmdGenerate.Name = "cmdGenerate";
            this.cmdGenerate.UseVisualStyleBackColor = true;
            this.cmdGenerate.Click += new System.EventHandler(this.cmdGenerate_Click);
            // 
            // lblTimeUnit
            // 
            this.lblTimeUnit.AccessibleDescription = null;
            this.lblTimeUnit.AccessibleName = null;
            resources.ApplyResources(this.lblTimeUnit, "lblTimeUnit");
            this.lblTimeUnit.Font = null;
            this.lblTimeUnit.Name = "lblTimeUnit";
            // 
            // nudActivateTime
            // 
            this.nudActivateTime.AccessibleDescription = null;
            this.nudActivateTime.AccessibleName = null;
            resources.ApplyResources(this.nudActivateTime, "nudActivateTime");
            this.nudActivateTime.Font = null;
            this.nudActivateTime.Maximum = new decimal(new int[] {
            365,
            0,
            0,
            0});
            this.nudActivateTime.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.nudActivateTime.Name = "nudActivateTime";
            this.nudActivateTime.Value = new decimal(new int[] {
            180,
            0,
            0,
            0});
            // 
            // lblActivateTime
            // 
            this.lblActivateTime.AccessibleDescription = null;
            this.lblActivateTime.AccessibleName = null;
            resources.ApplyResources(this.lblActivateTime, "lblActivateTime");
            this.lblActivateTime.Font = null;
            this.lblActivateTime.Name = "lblActivateTime";
            // 
            // ctlSaveFileDialog
            // 
            resources.ApplyResources(this.ctlSaveFileDialog, "ctlSaveFileDialog");
            // 
            // LicenseWF
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
            this.Name = "LicenseWF";
            this.pnlCover.ResumeLayout(false);
            this.pnlCover.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.nudActivateTime)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Label lblActivateTime;
        private System.Windows.Forms.NumericUpDown nudActivateTime;
        private System.Windows.Forms.Label lblTimeUnit;
        private System.Windows.Forms.Button cmdGenerate;
        private System.Windows.Forms.SaveFileDialog ctlSaveFileDialog;
    }
}