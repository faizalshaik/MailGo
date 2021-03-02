namespace DataDesign.MailGO.Accessories
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
            this.cmdGenerateActivation = new System.Windows.Forms.Button();
            this.cmdGenerateLicense = new System.Windows.Forms.Button();
            this.pnlCover.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlCover
            // 
            this.pnlCover.AccessibleDescription = null;
            this.pnlCover.AccessibleName = null;
            resources.ApplyResources(this.pnlCover, "pnlCover");
            this.pnlCover.BackgroundImage = null;
            this.pnlCover.Controls.Add(this.cmdGenerateActivation);
            this.pnlCover.Controls.Add(this.cmdGenerateLicense);
            this.pnlCover.Font = null;
            this.pnlCover.Name = "pnlCover";
            // 
            // cmdGenerateActivation
            // 
            this.cmdGenerateActivation.AccessibleDescription = null;
            this.cmdGenerateActivation.AccessibleName = null;
            resources.ApplyResources(this.cmdGenerateActivation, "cmdGenerateActivation");
            this.cmdGenerateActivation.BackgroundImage = null;
            this.cmdGenerateActivation.Font = null;
            this.cmdGenerateActivation.Name = "cmdGenerateActivation";
            this.cmdGenerateActivation.UseVisualStyleBackColor = true;
            this.cmdGenerateActivation.Click += new System.EventHandler(this.cmdGenerateActivation_Click);
            // 
            // cmdGenerateLicense
            // 
            this.cmdGenerateLicense.AccessibleDescription = null;
            this.cmdGenerateLicense.AccessibleName = null;
            resources.ApplyResources(this.cmdGenerateLicense, "cmdGenerateLicense");
            this.cmdGenerateLicense.BackgroundImage = null;
            this.cmdGenerateLicense.Font = null;
            this.cmdGenerateLicense.Name = "cmdGenerateLicense";
            this.cmdGenerateLicense.UseVisualStyleBackColor = true;
            this.cmdGenerateLicense.Click += new System.EventHandler(this.cmdGenerateLicense_Click);
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
            this.pnlCover.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Button cmdGenerateActivation;
        private System.Windows.Forms.Button cmdGenerateLicense;
    }
}