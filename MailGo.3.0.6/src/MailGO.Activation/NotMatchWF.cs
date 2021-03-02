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

namespace DataDesign.MailGO.Activation
{
    internal partial class NotMatchWF : Form
    {
        private string m_fileName;
        private string m_fileProperty;
        private string m_emailAddress;
        private bool m_closed = false;

        public NotMatchWF(string v_fileName, string v_fileProperty, string v_emailAddress)
        {
            this.m_fileName = v_fileName;
            this.m_fileProperty = v_fileProperty;
            this.m_emailAddress = v_emailAddress;
            InitializeComponent();
            this.TopMost = true;
        }

        private void NotMatchWF_Load(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.None;
            this.lblFileName.Text += " " + this.m_fileName;
            //this.lblFileProperty.Text += " " + this.m_fileProperty;
            lblProperty.Text = this.m_fileProperty;
            this.lblAddress.Text += " " + this.m_emailAddress;
        }

        private void cmdSend_Click(object sender, EventArgs e)
        {
            this.m_closed = true;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.m_closed = true;
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void NotMatchWF_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!this.m_closed)
            {
                e.Cancel = true;
            }
        }

    }
}