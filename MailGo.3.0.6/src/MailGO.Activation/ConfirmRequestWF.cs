/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.02.18 Tran Dinh Thoai
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
    internal partial class ConfirmRequestWF : Form
    {
        public ConfirmRequestWF()
        {
            InitializeComponent();
        }

        private void cmdYes_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Yes;
            this.Close();
        }

        private void cmdNo_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
            this.Close();
        }

        private void ConfirmRequestWF_Load(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.None;
        }
    }
}