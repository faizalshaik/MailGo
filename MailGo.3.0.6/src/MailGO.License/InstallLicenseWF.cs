/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.24 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace DataDesign.MailGO.License
{
    internal partial class InstallLicenseWF : Form
    {
        private Model.IMailGoPG m_mailgo;

        public InstallLicenseWF(Model.IMailGoPG v_mailgo)
        {
            this.m_mailgo = v_mailgo;
            InitializeComponent();
        }

        private void cmdOpen_Click(object sender, EventArgs e)
        {
            this.ctlOpenFileDialog.ShowDialog(this);

            if (this.ctlOpenFileDialog.FileName == "") return;

            string t_license_id = File.ReadAllText(this.ctlOpenFileDialog.FileName);
            this.txtLicenseID.Text = t_license_id;
        }

        private void cmdInstall_Click(object sender, EventArgs e)
        {
            try
            {
                this.m_mailgo.License.InstallLicense(this.txtLicenseID.Text);
                this.Close();
            }
            catch (Exception ex)
            {
                this.m_mailgo.Track.Error(ex);
                this.ShowMessage(this.lblMsgInvalidLicense);
            }
        }

        private void ShowMessage(Label v_message)
        {
            MessageBox.Show(v_message.Text, this.lblMsgError.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}