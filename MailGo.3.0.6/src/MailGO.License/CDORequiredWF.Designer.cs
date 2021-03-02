namespace DataDesign.MailGO.License
{
    partial class CDORequiredWF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CDORequiredWF));
            this.lblMsg2K7 = new System.Windows.Forms.Label();
            this.lblMsg2K3 = new System.Windows.Forms.Label();
            this.cmdOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblMsg2K7
            // 
            this.lblMsg2K7.AccessibleDescription = null;
            this.lblMsg2K7.AccessibleName = null;
            resources.ApplyResources(this.lblMsg2K7, "lblMsg2K7");
            this.lblMsg2K7.Font = null;
            this.lblMsg2K7.Name = "lblMsg2K7";
            // 
            // lblMsg2K3
            // 
            this.lblMsg2K3.AccessibleDescription = null;
            this.lblMsg2K3.AccessibleName = null;
            resources.ApplyResources(this.lblMsg2K3, "lblMsg2K3");
            this.lblMsg2K3.Font = null;
            this.lblMsg2K3.Name = "lblMsg2K3";
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
            // CDORequiredWF
            // 
            this.AccessibleDescription = null;
            this.AccessibleName = null;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = null;
            this.Controls.Add(this.cmdOK);
            this.Controls.Add(this.lblMsg2K3);
            this.Controls.Add(this.lblMsg2K7);
            this.Font = null;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CDORequiredWF";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblMsg2K7;
        private System.Windows.Forms.Label lblMsg2K3;
        private System.Windows.Forms.Button cmdOK;
    }
}