/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.23 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DataDesign.MailGO.Activation
{
    internal partial class OptionWF : Form
    {
        private Model.IMailGoPG m_mailgo;

        public OptionWF(Model.IMailGoPG v_mailgo)
        {
            this.m_mailgo = v_mailgo;
            InitializeComponent();

            cmbLanguage.Items.Add("English");
            cmbLanguage.Items.Add("日本語");

            if (m_mailgo.Activation.Language == "ja-JP")
                cmbLanguage.SelectedIndex = 1;
            else
                cmbLanguage.SelectedIndex = 0;
        }

        private void cmdCreateList_Click(object sender, EventArgs e)
        {
            this.m_mailgo.Address.CreateList();
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            this.m_mailgo.Activation.Option.CheckAllAddress = this.optCheckAllTO.Checked;
            this.m_mailgo.Activation.Option.CheckOnlyFirstAddress = this.optCheckOneTO.Checked;
            this.m_mailgo.Activation.Option.CheckCCLine = this.chkCheckCCLine.Checked;
            this.m_mailgo.Activation.Option.CheckSameSuffix = !this.chkNotCheckSame.Checked;
            // 2010/9/18
            this.m_mailgo.Activation.Option.CheckWord = this.checkWord.Checked;
            this.m_mailgo.Activation.Option.CheckExcel = this.checkExcel.Checked;
            this.m_mailgo.Activation.Option.CheckPowerPoint = this.checkPowerPoint.Checked;
            this.m_mailgo.Activation.Option.CheckText = this.checkText.Checked;
            //
            this.m_mailgo.Activation.SaveOption();
            this.Close();
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void OptionWF_Load(object sender, EventArgs e)
        {
            this.tmrLoad.Enabled = true;
        }

        private void cmdActivate_Click(object sender, EventArgs e)
        {
            this.m_mailgo.Activation.Activate();
            this.Close();
        }

        private void HideDeadLine()
        {
            //if (this.cmdActivate.Visible)
            //{
            //    int t_gap = this.lnLable0.Top - this.lblLine06.Top;
            //    this.lblLine06.Visible = false;
            //    this.lblActivationDeadline.Visible = false;
            //    this.lblActivationDeadlineMsg.Visible = false;
            //    this.cmdActivate.Visible = false;
            //    this.Height -= t_gap;
            //}
        }

        private void tmrLoad_Tick(object sender, EventArgs e)
        {
            this.tmrLoad.Enabled = false;

            if(m_mailgo.Activation.ActivatedReally)
            {
                //this.HideDeadLine();
                this.cmdActivate.Visible = false;
                this.lblActivationDeadlineMsg.Visible = false;
                this.cmdDeActivate.Visible = true;
            }
            else
            {
                this.lblActivationDeadlineMsg.Text = string.Format(this.lblActivationDeadlineMsg.Text, m_mailgo.Activation.DeadLine.ToString("yyyy/MM/dd"));
                this.lblActivationDeadlineMsg.Visible = true;
                this.cmdActivate.Visible = true;
                this.cmdDeActivate.Visible = false;
            }

            this.optCheckAllTO.Checked = this.m_mailgo.Activation.Option.CheckAllAddress;
            this.optCheckOneTO.Checked = this.m_mailgo.Activation.Option.CheckOnlyFirstAddress;
            this.chkCheckCCLine.Checked = this.m_mailgo.Activation.Option.CheckCCLine;
            this.chkNotCheckSame.Checked = !this.m_mailgo.Activation.Option.CheckSameSuffix;
            // 2010/9/15
            this.checkWord.Checked = this.m_mailgo.Activation.Option.CheckWord;
            this.checkExcel.Checked = this.m_mailgo.Activation.Option.CheckExcel;
            this.checkPowerPoint.Checked = this.m_mailgo.Activation.Option.CheckPowerPoint;
            this.checkText.Checked = this.m_mailgo.Activation.Option.CheckText;
            //
        }

        private void lblUserDefinedList_Click(object sender, EventArgs e)
        {

        }

        private void lblUserDefinedInfo_Click(object sender, EventArgs e)
        {

        }

        private void pnlCover_Paint(object sender, PaintEventArgs e)
        {

        }

        private void lblFilePropertyOption_Click(object sender, EventArgs e)
        {

        }

        private void lblLine07_Click(object sender, EventArgs e)
        {

        }

        private void chkNotCheckSame_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void lblSuffixOption_Click(object sender, EventArgs e)
        {

        }

        private void chkCheckCCLine_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void lblCCLineOption_Click(object sender, EventArgs e)
        {

        }

        private void optCheckOneTO_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void optCheckAllTO_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void lblTOLineOption_Click(object sender, EventArgs e)
        {

        }

        private void lblLine03_Click(object sender, EventArgs e)
        {

        }

        private void lblLine02_Click(object sender, EventArgs e)
        {

        }

        private void lblLine04_Click(object sender, EventArgs e)
        {

        }

        private void lblLine01_Click(object sender, EventArgs e)
        {

        }

        private void checkWord_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkExcel_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkPowerPoint_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkText_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btnChangeLanguage_Click(object sender, EventArgs e)
        {
            string newLang = "en-US";
            if (cmbLanguage.SelectedIndex == 1)
                newLang = "ja-JP";

            if (m_mailgo.Activation.Language == newLang)
                return;

            m_mailgo.Activation.Language = newLang;
            Close();
        }

        private void cmdDeActivate_Click(object sender, EventArgs e)
        {
            this.m_mailgo.Activation.DeActivate();
            this.Close();
        }
    }
}