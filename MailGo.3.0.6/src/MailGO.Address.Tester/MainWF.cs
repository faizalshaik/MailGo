/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.18 Tran Dinh Thoai
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
    internal partial class MainWF : Form
    {
        private Model.IMailGoPG m_mailgo = new MailGoPG();

        public MainWF()
        {
            InitializeComponent();
        }

        private void MainWF_Load(object sender, EventArgs e)
        {
            this.cboLanguage.SelectedIndex = 0;
            this.m_mailgo.Startup();
        }

        private void MainWF_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.m_mailgo.Shutdown();
        }

        private void cmdCreateList_Click(object sender, EventArgs e)
        {
            this.m_mailgo.Address.CreateList();
        }

        private void cmdCheckEmail_Click(object sender, EventArgs e)
        {
            CheckerWF t_checker = new CheckerWF(this.m_mailgo);
            t_checker.ShowDialog(this);
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

    }
}