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
using System.Windows.Forms;

namespace DataDesign.MailGO.Accessories
{
    internal partial class MainWF : Form
    {
        private Model.IMailGoPG m_mailgo = new MailGoPG();

        public MainWF()
        {
            InitializeComponent();
        }

        private void cmdGenerateLicense_Click(object sender, EventArgs e)
        {
            this.m_mailgo.License.GenerateLicenseID();
        }

        private void cmdGenerateActivation_Click(object sender, EventArgs e)
        {
            this.m_mailgo.License.GenerateActivationID();
        }
    }
}