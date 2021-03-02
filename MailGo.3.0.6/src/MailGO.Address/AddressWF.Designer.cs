namespace DataDesign.MailGO.Address
{
    partial class AddressWF
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddressWF));
            this.pnlCover = new System.Windows.Forms.Panel();
            this.pnlUserDefined = new System.Windows.Forms.Panel();
            this.pnlAction = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.cmdOK = new System.Windows.Forms.Button();
            this.pnlEditor = new System.Windows.Forms.Panel();
            this.cmdExport = new System.Windows.Forms.Button();
            this.cmdImport = new System.Windows.Forms.Button();
            this.cmdDelete = new System.Windows.Forms.Button();
            this.cmdModify = new System.Windows.Forms.Button();
            this.cmdNew = new System.Windows.Forms.Button();
            this.txtSuffix = new System.Windows.Forms.TextBox();
            this.lblSuffix = new System.Windows.Forms.Label();
            this.txtCompanyName = new System.Windows.Forms.TextBox();
            this.lblCompanyName = new System.Windows.Forms.Label();
            this.cmdSave = new System.Windows.Forms.Button();
            this.cmdDiscard = new System.Windows.Forms.Button();
            this.lsvUserDefined = new System.Windows.Forms.ListView();
            this.colCompanyU = new System.Windows.Forms.ColumnHeader();
            this.colSuffixU = new System.Windows.Forms.ColumnHeader();
            this.lblUserDefined = new System.Windows.Forms.Label();
            this.ctlOpenFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.ctlSaveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.pnlCover.SuspendLayout();
            this.pnlUserDefined.SuspendLayout();
            this.pnlAction.SuspendLayout();
            this.pnlEditor.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlCover
            // 
            this.pnlCover.Controls.Add(this.pnlUserDefined);
            resources.ApplyResources(this.pnlCover, "pnlCover");
            this.pnlCover.Name = "pnlCover";
            // 
            // pnlUserDefined
            // 
            this.pnlUserDefined.Controls.Add(this.pnlAction);
            this.pnlUserDefined.Controls.Add(this.pnlEditor);
            this.pnlUserDefined.Controls.Add(this.lsvUserDefined);
            this.pnlUserDefined.Controls.Add(this.lblUserDefined);
            resources.ApplyResources(this.pnlUserDefined, "pnlUserDefined");
            this.pnlUserDefined.Name = "pnlUserDefined";
            // 
            // pnlAction
            // 
            this.pnlAction.Controls.Add(this.label1);
            this.pnlAction.Controls.Add(this.cmdCancel);
            this.pnlAction.Controls.Add(this.cmdOK);
            resources.ApplyResources(this.pnlAction, "pnlAction");
            this.pnlAction.Name = "pnlAction";
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // cmdCancel
            // 
            resources.ApplyResources(this.cmdCancel, "cmdCancel");
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.UseVisualStyleBackColor = true;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // cmdOK
            // 
            resources.ApplyResources(this.cmdOK, "cmdOK");
            this.cmdOK.Name = "cmdOK";
            this.cmdOK.UseVisualStyleBackColor = true;
            this.cmdOK.Click += new System.EventHandler(this.cmdOK_Click);
            // 
            // pnlEditor
            // 
            this.pnlEditor.Controls.Add(this.cmdExport);
            this.pnlEditor.Controls.Add(this.cmdImport);
            this.pnlEditor.Controls.Add(this.cmdDelete);
            this.pnlEditor.Controls.Add(this.cmdModify);
            this.pnlEditor.Controls.Add(this.cmdNew);
            this.pnlEditor.Controls.Add(this.txtSuffix);
            this.pnlEditor.Controls.Add(this.lblSuffix);
            this.pnlEditor.Controls.Add(this.txtCompanyName);
            this.pnlEditor.Controls.Add(this.lblCompanyName);
            this.pnlEditor.Controls.Add(this.cmdSave);
            this.pnlEditor.Controls.Add(this.cmdDiscard);
            resources.ApplyResources(this.pnlEditor, "pnlEditor");
            this.pnlEditor.Name = "pnlEditor";
            // 
            // cmdExport
            // 
            resources.ApplyResources(this.cmdExport, "cmdExport");
            this.cmdExport.Name = "cmdExport";
            this.cmdExport.UseVisualStyleBackColor = true;
            this.cmdExport.Click += new System.EventHandler(this.cmdExport_Click);
            // 
            // cmdImport
            // 
            resources.ApplyResources(this.cmdImport, "cmdImport");
            this.cmdImport.Name = "cmdImport";
            this.cmdImport.UseVisualStyleBackColor = true;
            this.cmdImport.Click += new System.EventHandler(this.cmdImport_Click);
            // 
            // cmdDelete
            // 
            resources.ApplyResources(this.cmdDelete, "cmdDelete");
            this.cmdDelete.Name = "cmdDelete";
            this.cmdDelete.UseVisualStyleBackColor = true;
            this.cmdDelete.Click += new System.EventHandler(this.cmdDelete_Click);
            // 
            // cmdModify
            // 
            resources.ApplyResources(this.cmdModify, "cmdModify");
            this.cmdModify.Name = "cmdModify";
            this.cmdModify.UseVisualStyleBackColor = true;
            this.cmdModify.Click += new System.EventHandler(this.cmdModify_Click);
            // 
            // cmdNew
            // 
            resources.ApplyResources(this.cmdNew, "cmdNew");
            this.cmdNew.Name = "cmdNew";
            this.cmdNew.UseVisualStyleBackColor = true;
            this.cmdNew.Click += new System.EventHandler(this.cmdNew_Click);
            // 
            // txtSuffix
            // 
            resources.ApplyResources(this.txtSuffix, "txtSuffix");
            this.txtSuffix.Name = "txtSuffix";
            // 
            // lblSuffix
            // 
            resources.ApplyResources(this.lblSuffix, "lblSuffix");
            this.lblSuffix.Name = "lblSuffix";
            // 
            // txtCompanyName
            // 
            resources.ApplyResources(this.txtCompanyName, "txtCompanyName");
            this.txtCompanyName.Name = "txtCompanyName";
            // 
            // lblCompanyName
            // 
            resources.ApplyResources(this.lblCompanyName, "lblCompanyName");
            this.lblCompanyName.Name = "lblCompanyName";
            // 
            // cmdSave
            // 
            resources.ApplyResources(this.cmdSave, "cmdSave");
            this.cmdSave.Name = "cmdSave";
            this.cmdSave.UseVisualStyleBackColor = true;
            this.cmdSave.Click += new System.EventHandler(this.cmdSave_Click);
            // 
            // cmdDiscard
            // 
            resources.ApplyResources(this.cmdDiscard, "cmdDiscard");
            this.cmdDiscard.Name = "cmdDiscard";
            this.cmdDiscard.UseVisualStyleBackColor = true;
            this.cmdDiscard.Click += new System.EventHandler(this.cmdDiscard_Click);
            // 
            // lsvUserDefined
            // 
            resources.ApplyResources(this.lsvUserDefined, "lsvUserDefined");
            this.lsvUserDefined.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colCompanyU,
            this.colSuffixU});
            this.lsvUserDefined.FullRowSelect = true;
            this.lsvUserDefined.MultiSelect = false;
            this.lsvUserDefined.Name = "lsvUserDefined";
            this.lsvUserDefined.UseCompatibleStateImageBehavior = false;
            this.lsvUserDefined.View = System.Windows.Forms.View.Details;
            this.lsvUserDefined.SelectedIndexChanged += new System.EventHandler(this.lsvUserDefined_SelectedIndexChanged);
            // 
            // colCompanyU
            // 
            resources.ApplyResources(this.colCompanyU, "colCompanyU");
            // 
            // colSuffixU
            // 
            resources.ApplyResources(this.colSuffixU, "colSuffixU");
            // 
            // lblUserDefined
            // 
            resources.ApplyResources(this.lblUserDefined, "lblUserDefined");
            this.lblUserDefined.Name = "lblUserDefined";
            // 
            // ctlOpenFileDialog
            // 
            resources.ApplyResources(this.ctlOpenFileDialog, "ctlOpenFileDialog");
            // 
            // AddressWF
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pnlCover);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AddressWF";
            this.Load += new System.EventHandler(this.AddressWF_Load);
            this.pnlCover.ResumeLayout(false);
            this.pnlUserDefined.ResumeLayout(false);
            this.pnlUserDefined.PerformLayout();
            this.pnlAction.ResumeLayout(false);
            this.pnlEditor.ResumeLayout(false);
            this.pnlEditor.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Panel pnlUserDefined;
        private System.Windows.Forms.ListView lsvUserDefined;
        private System.Windows.Forms.ColumnHeader colCompanyU;
        private System.Windows.Forms.ColumnHeader colSuffixU;
        private System.Windows.Forms.Label lblUserDefined;
        private System.Windows.Forms.Panel pnlEditor;
        private System.Windows.Forms.Label lblCompanyName;
        private System.Windows.Forms.TextBox txtCompanyName;
        private System.Windows.Forms.TextBox txtSuffix;
        private System.Windows.Forms.Label lblSuffix;
        private System.Windows.Forms.Button cmdNew;
        private System.Windows.Forms.Button cmdModify;
        private System.Windows.Forms.Button cmdSave;
        private System.Windows.Forms.Button cmdDelete;
        private System.Windows.Forms.Button cmdDiscard;
        private System.Windows.Forms.Panel pnlAction;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.Button cmdOK;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button cmdImport;
        private System.Windows.Forms.OpenFileDialog ctlOpenFileDialog;
        private System.Windows.Forms.Button cmdExport;
        private System.Windows.Forms.SaveFileDialog ctlSaveFileDialog;
    }
}