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
using System.Windows.Forms;

namespace DataDesign.MailGO.License.Tester
{
    public partial class RequestWF : Form
    {
        private Model.ILicensePG m_license;

        public RequestWF(Model.ILicensePG v_license)
        {
            this.m_license = v_license;
            InitializeComponent();
        }

        private void cmdEncode_Click(object sender, EventArgs e)
        {
            string t_request;
            this.m_license.RequestActivationID(out t_request, this.txtLicenseId.Text, this.LoadHardware());
            this.txtRequest.Text = t_request;
        }

        private void cmdDecode_Click(object sender, EventArgs e)
        {
            Model.IHardware t_hardware;
            string t_license_id;
            this.m_license.RequestActivationID(out t_license_id, out t_hardware, this.txtRequest.Text);
            this.txtLicenseId.Text = t_license_id;
            this.DisplayHardware(t_hardware);
        }

        private void DisplayHardware(Model.IHardware v_hardware)
        {
            this.txtMacList.Text = string.Join(Environment.NewLine, v_hardware.MAC.ToArray());
            this.txtCpuList.Text = string.Join(Environment.NewLine, v_hardware.CPU.ToArray());
        }

        private Model.IHardware LoadHardware()
        {
            Model.IHardware t_hardware = new Model.MHardware();

            foreach (string t_mac in this.txtMacList.Lines)
            {
                t_hardware.MAC.Add(t_mac);
            }

            foreach (string t_cpu in this.txtCpuList.Lines)
            {
                t_hardware.CPU.Add(t_cpu);
            }

            return t_hardware;
        }
    }
}