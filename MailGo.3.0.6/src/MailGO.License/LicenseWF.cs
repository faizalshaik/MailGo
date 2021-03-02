/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.22 Tran Dinh Thoai
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
    internal partial class LicenseWF : Form
    {
        private Model.ILicensePG m_license;

        public LicenseWF(Model.ILicensePG v_license)
        {
            this.m_license = v_license;
            InitializeComponent();
        }

        private void cmdGenerate_Click(object sender, EventArgs e)
        {
            this.ctlSaveFileDialog.FileName = "";
            this.ctlSaveFileDialog.ShowDialog(this);

            if (this.ctlSaveFileDialog.FileName == "") return;

            string t_license = this.m_license.GenerateLicenseID(false, (int)this.nudActivateTime.Value);

            StreamWriter t_writer = new StreamWriter(File.Open(this.ctlSaveFileDialog.FileName, FileMode.Create, FileAccess.Write));
            t_writer.Write(t_license);
            t_writer.Close();
        }
    }
}