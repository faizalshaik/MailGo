namespace DataDesign.MailGO.Activation
{
    partial class OptionWF
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OptionWF));
            this.pnlCover = new System.Windows.Forms.Panel();
            this.cmdDeActivate = new System.Windows.Forms.Button();
            this.btnChangeLanguage = new System.Windows.Forms.Button();
            this.cmbLanguage = new System.Windows.Forms.ComboBox();
            this.checkText = new System.Windows.Forms.CheckBox();
            this.checkPowerPoint = new System.Windows.Forms.CheckBox();
            this.checkExcel = new System.Windows.Forms.CheckBox();
            this.checkWord = new System.Windows.Forms.CheckBox();
            this.cmdActivate = new System.Windows.Forms.Button();
            this.lblActivationDeadlineMsg = new System.Windows.Forms.Label();
            this.lblActivationDeadline = new System.Windows.Forms.Label();
            this.lblLine06 = new System.Windows.Forms.Label();
            this.lblInfo = new System.Windows.Forms.Label();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.cmdOK = new System.Windows.Forms.Button();
            this.chkNotCheckSame = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lblSuffixOption = new System.Windows.Forms.Label();
            this.chkCheckCCLine = new System.Windows.Forms.CheckBox();
            this.lblCCLineOption = new System.Windows.Forms.Label();
            this.optCheckOneTO = new System.Windows.Forms.RadioButton();
            this.optCheckAllTO = new System.Windows.Forms.RadioButton();
            this.lblTOLineOption = new System.Windows.Forms.Label();
            this.cmdCreateList = new System.Windows.Forms.Button();
            this.lblUserDefinedInfo = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.lblUserDefinedList = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblLine03 = new System.Windows.Forms.Label();
            this.lblLine02 = new System.Windows.Forms.Label();
            this.lnLable0 = new System.Windows.Forms.Label();
            this.lblLine04 = new System.Windows.Forms.Label();
            this.lblLine01 = new System.Windows.Forms.Label();
            this.tmrLoad = new System.Windows.Forms.Timer(this.components);
            this.pnlCover.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlCover
            // 
            resources.ApplyResources(this.pnlCover, "pnlCover");
            this.pnlCover.Controls.Add(this.cmdDeActivate);
            this.pnlCover.Controls.Add(this.btnChangeLanguage);
            this.pnlCover.Controls.Add(this.cmbLanguage);
            this.pnlCover.Controls.Add(this.checkText);
            this.pnlCover.Controls.Add(this.checkPowerPoint);
            this.pnlCover.Controls.Add(this.checkExcel);
            this.pnlCover.Controls.Add(this.checkWord);
            this.pnlCover.Controls.Add(this.cmdActivate);
            this.pnlCover.Controls.Add(this.lblActivationDeadlineMsg);
            this.pnlCover.Controls.Add(this.lblActivationDeadline);
            this.pnlCover.Controls.Add(this.lblLine06);
            this.pnlCover.Controls.Add(this.lblInfo);
            this.pnlCover.Controls.Add(this.cmdCancel);
            this.pnlCover.Controls.Add(this.cmdOK);
            this.pnlCover.Controls.Add(this.chkNotCheckSame);
            this.pnlCover.Controls.Add(this.label3);
            this.pnlCover.Controls.Add(this.lblSuffixOption);
            this.pnlCover.Controls.Add(this.chkCheckCCLine);
            this.pnlCover.Controls.Add(this.lblCCLineOption);
            this.pnlCover.Controls.Add(this.optCheckOneTO);
            this.pnlCover.Controls.Add(this.optCheckAllTO);
            this.pnlCover.Controls.Add(this.lblTOLineOption);
            this.pnlCover.Controls.Add(this.cmdCreateList);
            this.pnlCover.Controls.Add(this.lblUserDefinedInfo);
            this.pnlCover.Controls.Add(this.label2);
            this.pnlCover.Controls.Add(this.lblUserDefinedList);
            this.pnlCover.Controls.Add(this.label4);
            this.pnlCover.Controls.Add(this.lblLine03);
            this.pnlCover.Controls.Add(this.lblLine02);
            this.pnlCover.Controls.Add(this.lnLable0);
            this.pnlCover.Controls.Add(this.lblLine04);
            this.pnlCover.Controls.Add(this.lblLine01);
            this.pnlCover.Name = "pnlCover";
            this.pnlCover.Paint += new System.Windows.Forms.PaintEventHandler(this.pnlCover_Paint);
            // 
            // cmdDeActivate
            // 
            resources.ApplyResources(this.cmdDeActivate, "cmdDeActivate");
            this.cmdDeActivate.Name = "cmdDeActivate";
            this.cmdDeActivate.UseVisualStyleBackColor = true;
            this.cmdDeActivate.Click += new System.EventHandler(this.cmdDeActivate_Click);
            // 
            // btnChangeLanguage
            // 
            resources.ApplyResources(this.btnChangeLanguage, "btnChangeLanguage");
            this.btnChangeLanguage.Name = "btnChangeLanguage";
            this.btnChangeLanguage.UseVisualStyleBackColor = true;
            this.btnChangeLanguage.Click += new System.EventHandler(this.btnChangeLanguage_Click);
            // 
            // cmbLanguage
            // 
            resources.ApplyResources(this.cmbLanguage, "cmbLanguage");
            this.cmbLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbLanguage.FormattingEnabled = true;
            this.cmbLanguage.Name = "cmbLanguage";
            // 
            // checkText
            // 
            resources.ApplyResources(this.checkText, "checkText");
            this.checkText.Name = "checkText";
            this.checkText.UseVisualStyleBackColor = true;
            this.checkText.CheckedChanged += new System.EventHandler(this.checkText_CheckedChanged);
            // 
            // checkPowerPoint
            // 
            resources.ApplyResources(this.checkPowerPoint, "checkPowerPoint");
            this.checkPowerPoint.Name = "checkPowerPoint";
            this.checkPowerPoint.UseVisualStyleBackColor = true;
            this.checkPowerPoint.CheckedChanged += new System.EventHandler(this.checkPowerPoint_CheckedChanged);
            // 
            // checkExcel
            // 
            resources.ApplyResources(this.checkExcel, "checkExcel");
            this.checkExcel.Name = "checkExcel";
            this.checkExcel.UseVisualStyleBackColor = true;
            this.checkExcel.CheckedChanged += new System.EventHandler(this.checkExcel_CheckedChanged);
            // 
            // checkWord
            // 
            resources.ApplyResources(this.checkWord, "checkWord");
            this.checkWord.Name = "checkWord";
            this.checkWord.UseVisualStyleBackColor = true;
            this.checkWord.CheckedChanged += new System.EventHandler(this.checkWord_CheckedChanged);
            // 
            // cmdActivate
            // 
            resources.ApplyResources(this.cmdActivate, "cmdActivate");
            this.cmdActivate.Name = "cmdActivate";
            this.cmdActivate.UseVisualStyleBackColor = true;
            this.cmdActivate.Click += new System.EventHandler(this.cmdActivate_Click);
            // 
            // lblActivationDeadlineMsg
            // 
            resources.ApplyResources(this.lblActivationDeadlineMsg, "lblActivationDeadlineMsg");
            this.lblActivationDeadlineMsg.Name = "lblActivationDeadlineMsg";
            // 
            // lblActivationDeadline
            // 
            resources.ApplyResources(this.lblActivationDeadline, "lblActivationDeadline");
            this.lblActivationDeadline.Name = "lblActivationDeadline";
            // 
            // lblLine06
            // 
            resources.ApplyResources(this.lblLine06, "lblLine06");
            this.lblLine06.BackColor = System.Drawing.SystemColors.Control;
            this.lblLine06.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLine06.Name = "lblLine06";
            // 
            // lblInfo
            // 
            resources.ApplyResources(this.lblInfo, "lblInfo");
            this.lblInfo.Name = "lblInfo";
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
            // chkNotCheckSame
            // 
            resources.ApplyResources(this.chkNotCheckSame, "chkNotCheckSame");
            this.chkNotCheckSame.Name = "chkNotCheckSame";
            this.chkNotCheckSame.UseVisualStyleBackColor = true;
            this.chkNotCheckSame.CheckedChanged += new System.EventHandler(this.chkNotCheckSame_CheckedChanged);
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            this.label3.Click += new System.EventHandler(this.lblSuffixOption_Click);
            // 
            // lblSuffixOption
            // 
            resources.ApplyResources(this.lblSuffixOption, "lblSuffixOption");
            this.lblSuffixOption.Name = "lblSuffixOption";
            this.lblSuffixOption.Click += new System.EventHandler(this.lblSuffixOption_Click);
            // 
            // chkCheckCCLine
            // 
            resources.ApplyResources(this.chkCheckCCLine, "chkCheckCCLine");
            this.chkCheckCCLine.Name = "chkCheckCCLine";
            this.chkCheckCCLine.UseVisualStyleBackColor = true;
            this.chkCheckCCLine.CheckedChanged += new System.EventHandler(this.chkCheckCCLine_CheckedChanged);
            // 
            // lblCCLineOption
            // 
            resources.ApplyResources(this.lblCCLineOption, "lblCCLineOption");
            this.lblCCLineOption.Name = "lblCCLineOption";
            this.lblCCLineOption.Click += new System.EventHandler(this.lblCCLineOption_Click);
            // 
            // optCheckOneTO
            // 
            resources.ApplyResources(this.optCheckOneTO, "optCheckOneTO");
            this.optCheckOneTO.Name = "optCheckOneTO";
            this.optCheckOneTO.UseVisualStyleBackColor = true;
            this.optCheckOneTO.CheckedChanged += new System.EventHandler(this.optCheckOneTO_CheckedChanged);
            // 
            // optCheckAllTO
            // 
            resources.ApplyResources(this.optCheckAllTO, "optCheckAllTO");
            this.optCheckAllTO.Checked = true;
            this.optCheckAllTO.Name = "optCheckAllTO";
            this.optCheckAllTO.TabStop = true;
            this.optCheckAllTO.UseVisualStyleBackColor = true;
            this.optCheckAllTO.CheckedChanged += new System.EventHandler(this.optCheckAllTO_CheckedChanged);
            // 
            // lblTOLineOption
            // 
            resources.ApplyResources(this.lblTOLineOption, "lblTOLineOption");
            this.lblTOLineOption.Name = "lblTOLineOption";
            this.lblTOLineOption.Click += new System.EventHandler(this.lblTOLineOption_Click);
            // 
            // cmdCreateList
            // 
            resources.ApplyResources(this.cmdCreateList, "cmdCreateList");
            this.cmdCreateList.Name = "cmdCreateList";
            this.cmdCreateList.UseVisualStyleBackColor = true;
            this.cmdCreateList.Click += new System.EventHandler(this.cmdCreateList_Click);
            // 
            // lblUserDefinedInfo
            // 
            resources.ApplyResources(this.lblUserDefinedInfo, "lblUserDefinedInfo");
            this.lblUserDefinedInfo.Name = "lblUserDefinedInfo";
            this.lblUserDefinedInfo.Click += new System.EventHandler(this.lblUserDefinedInfo_Click);
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            this.label2.Click += new System.EventHandler(this.lblUserDefinedList_Click);
            // 
            // lblUserDefinedList
            // 
            resources.ApplyResources(this.lblUserDefinedList, "lblUserDefinedList");
            this.lblUserDefinedList.Name = "lblUserDefinedList";
            this.lblUserDefinedList.Click += new System.EventHandler(this.lblUserDefinedList_Click);
            // 
            // label4
            // 
            resources.ApplyResources(this.label4, "label4");
            this.label4.BackColor = System.Drawing.SystemColors.Control;
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label4.Name = "label4";
            this.label4.Click += new System.EventHandler(this.lblLine03_Click);
            // 
            // lblLine03
            // 
            resources.ApplyResources(this.lblLine03, "lblLine03");
            this.lblLine03.BackColor = System.Drawing.SystemColors.Control;
            this.lblLine03.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLine03.Name = "lblLine03";
            this.lblLine03.Click += new System.EventHandler(this.lblLine03_Click);
            // 
            // lblLine02
            // 
            resources.ApplyResources(this.lblLine02, "lblLine02");
            this.lblLine02.BackColor = System.Drawing.SystemColors.Control;
            this.lblLine02.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLine02.Name = "lblLine02";
            this.lblLine02.Click += new System.EventHandler(this.lblLine02_Click);
            // 
            // lnLable0
            // 
            resources.ApplyResources(this.lnLable0, "lnLable0");
            this.lnLable0.BackColor = System.Drawing.SystemColors.Control;
            this.lnLable0.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lnLable0.Name = "lnLable0";
            this.lnLable0.Click += new System.EventHandler(this.lblLine01_Click);
            // 
            // lblLine04
            // 
            resources.ApplyResources(this.lblLine04, "lblLine04");
            this.lblLine04.BackColor = System.Drawing.SystemColors.Control;
            this.lblLine04.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLine04.Name = "lblLine04";
            this.lblLine04.Click += new System.EventHandler(this.lblLine04_Click);
            // 
            // lblLine01
            // 
            resources.ApplyResources(this.lblLine01, "lblLine01");
            this.lblLine01.BackColor = System.Drawing.SystemColors.Control;
            this.lblLine01.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLine01.Name = "lblLine01";
            this.lblLine01.Click += new System.EventHandler(this.lblLine01_Click);
            // 
            // tmrLoad
            // 
            this.tmrLoad.Interval = 10;
            this.tmrLoad.Tick += new System.EventHandler(this.tmrLoad_Tick);
            // 
            // OptionWF
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pnlCover);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "OptionWF";
            this.Load += new System.EventHandler(this.OptionWF_Load);
            this.pnlCover.ResumeLayout(false);
            this.pnlCover.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel pnlCover;
        private System.Windows.Forms.Label lblLine01;
        private System.Windows.Forms.Label lblLine03;
        private System.Windows.Forms.Label lblLine02;
        private System.Windows.Forms.Label lblLine04;
        private System.Windows.Forms.Label lblUserDefinedList;
        private System.Windows.Forms.Label lblUserDefinedInfo;
        private System.Windows.Forms.Button cmdCreateList;
        private System.Windows.Forms.Label lblTOLineOption;
        private System.Windows.Forms.RadioButton optCheckAllTO;
        private System.Windows.Forms.RadioButton optCheckOneTO;
        private System.Windows.Forms.Label lblCCLineOption;
        private System.Windows.Forms.CheckBox chkCheckCCLine;
        private System.Windows.Forms.Label lblSuffixOption;
        private System.Windows.Forms.CheckBox chkNotCheckSame;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.Button cmdOK;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.Button cmdActivate;
        private System.Windows.Forms.Label lblActivationDeadlineMsg;
        private System.Windows.Forms.Label lblActivationDeadline;
        private System.Windows.Forms.Label lblLine06;
        private System.Windows.Forms.Timer tmrLoad;
        private System.Windows.Forms.CheckBox checkWord;
        private System.Windows.Forms.CheckBox checkPowerPoint;
        private System.Windows.Forms.CheckBox checkExcel;
        private System.Windows.Forms.CheckBox checkText;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label lnLable0;
        private System.Windows.Forms.Button btnChangeLanguage;
        private System.Windows.Forms.ComboBox cmbLanguage;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button cmdDeActivate;
    }
}