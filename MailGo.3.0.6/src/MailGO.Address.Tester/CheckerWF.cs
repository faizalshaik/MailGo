/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.21 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DataDesign.MailGO.Address.Tester
{
    public partial class CheckerWF : Form
    {
        private Model.IMailGoPG m_mailgo;

        public CheckerWF(Model.IMailGoPG v_mailgo)
        {
            this.m_mailgo = v_mailgo;
            InitializeComponent();
        }

        private void cmdCheck_Click(object sender, EventArgs e)
        {
            this.m_mailgo.Activation.Option.CheckAllAddress = this.chkCheckAllToAddr.Checked;
            this.m_mailgo.Activation.Option.CheckCCLine = this.chkCheckCCLine.Checked;
            this.m_mailgo.Activation.Option.CheckSameSuffix = this.chkCheckSameSuffix.Checked;

            Model.IEmail t_email = new Model.MEmail();
            bool t_cancel = false;

            t_email.Sender = this.txtSender.Text;
            t_email.TO.AddRange(this.txtTo.Text.Split(new char[] { ';', ',' }));
            t_email.CC.AddRange(this.txtCC.Text.Split(new char[] { ';', ',' }));
            t_email.Body = this.txtBody.Text;

            this.m_mailgo.Address.CheckEmail(t_email, out t_cancel);

            if (t_cancel)
            {
                System.Windows.Forms.MessageBox.Show("Cancel");
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Send email");
            }
        }

    }
}