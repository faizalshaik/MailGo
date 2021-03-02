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
    internal partial class ActivationWF : Form
    {
        private Model.ILicensePG m_license;

        public ActivationWF(Model.ILicensePG v_license)
        {
            this.m_license = v_license;
            InitializeComponent();
        }

        private void cmdGenerate_Click(object sender, EventArgs e)
        {
            if (!CheckInput(this.txtRequest.Text)) return;

            this.ctlSaveFileDialog.FileName = "";
            this.ctlSaveFileDialog.ShowDialog(this);

            if (this.ctlSaveFileDialog.FileName == "") return;

            try
            {
                string t_activation;
                string t_license_id;
                Model.IHardware t_hardware;

                this.m_license.RequestActivationID(out t_license_id, out t_hardware, this.txtRequest.Text.Trim());
                t_activation = this.m_license.GenerateActivationID(t_license_id, t_hardware);

                StreamWriter t_writer = new StreamWriter(File.Open(this.ctlSaveFileDialog.FileName, FileMode.Create, FileAccess.Write));
                t_writer.Write(t_activation);
                t_writer.Close();
            }
            catch (Exception ex)
            {
                this.m_license.MailGo.Track.Error(ex);
                this.ShowMessage(this.lblMsgInvalidRequest);
            }
        }

        private bool CheckInput(string input)
        {
            if (input == null || input.Trim().Length == 0)
            {
                ShowWarning(this.lblMsgEmptyKey);
                return false;
            }
            input = input.Trim();
            int pos = input.IndexOf('|');
            if (pos < 0)
            {
                ShowWarning(this.lblMsgMissHardware);
                return false;
            }
            return true;
        }

        private void ShowWarning(Label v_message)
        {
            MessageBox.Show(v_message.Text, this.lblMsgWarn.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void ShowMessage(Label v_message)
        {
            MessageBox.Show(v_message.Text, this.lblMsgError.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}