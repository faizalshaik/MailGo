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
using System.IO;
using System.Windows.Forms;

namespace DataDesign.MailGO.License.Tester
{
    public partial class MainWF : Form
    {
        private Model.IMailGoPG m_mailgo = new MailGoPG();

        public MainWF()
        {
            InitializeComponent();
        }

        private void cmdReadHardware_Click(object sender, EventArgs e)
        {
            HardwareWF t_form = new HardwareWF(this.m_mailgo.License);
            t_form.ShowDialog(this);
        }

        private void cmdGenerateLicense_Click(object sender, EventArgs e)
        {
            this.m_mailgo.License.GenerateLicenseID();
        }

        private void cmdRequestActivation_Click(object sender, EventArgs e)
        {
            RequestWF t_form = new RequestWF(this.m_mailgo.License);
            t_form.ShowDialog(this);
        }

        private void cmdGenerateActivation_Click(object sender, EventArgs e)
        {
            this.m_mailgo.License.GenerateActivationID();
        }

        private void cmdInstallLicense_Click(object sender, EventArgs e)
        {
            this.m_mailgo.License.InstallLicense();
        }


        private void cmdActivateLicense_Click(object sender, EventArgs e)
        {
            this.m_mailgo.License.ActivateLicense();
        }

        private void cboLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.cboLanguage.SelectedIndex == 0)
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
            }
            if (this.cboLanguage.SelectedIndex == 1)
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ja-JP");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("ja-JP");
            }
        }

        private void MainWF_Load(object sender, EventArgs e)
        {
            this.cboLanguage.SelectedIndex = 0;
        }
    }
}