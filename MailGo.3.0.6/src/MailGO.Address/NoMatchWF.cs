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

namespace DataDesign.MailGO.Address
{
    internal partial class NoMatchWF : Form
    {
        private string m_company;
        private string m_email;
        private bool m_closed = false;

        public NoMatchWF(string v_company, string v_email)
        {
            this.m_company = v_company;
            this.m_email = v_email;
            InitializeComponent();
            this.TopMost = true;
        }

        private void NoMatchWF_Load(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.None;
            this.lblCompany.Text += " " + this.m_company;
            this.lblEmail.Text += " " + this.m_email;
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.m_closed = true;
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void cmdAddList_Click(object sender, EventArgs e)
        {
            this.m_closed = true;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void cmdIgnore_Click(object sender, EventArgs e)
        {
            this.m_closed = true;
            this.DialogResult = DialogResult.Ignore;
            this.Close();
        }

        private void NoMatchWF_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!this.m_closed)
            {
                e.Cancel = true;
            }
        }

    }
}