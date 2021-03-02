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
    internal partial class DeadlineWF : Form
    {
        private DateTime m_expired_date;

        public DeadlineWF(DateTime v_expired_date)
        {
            this.m_expired_date = v_expired_date;
            InitializeComponent();
        }

        private void DeadlineWF_Load(object sender, EventArgs e)
        {
            this.lblMessage.Text = string.Format(this.lblMessage.Text, this.m_expired_date.ToString("yyyy/MM/dd"));
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}