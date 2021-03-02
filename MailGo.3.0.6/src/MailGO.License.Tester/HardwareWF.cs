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

namespace DataDesign.MailGO.License.Tester
{
    public partial class HardwareWF : Form
    {
        private Model.ILicensePG m_license;

        public HardwareWF(Model.ILicensePG v_license)
        {
            this.m_license = v_license;
            InitializeComponent();
        }

        private void HardwareWF_Load(object sender, EventArgs e)
        {
            Model.IHardware t_hardware = this.m_license.ReadHardware();

            foreach (string t_mac in t_hardware.MAC)
            {
                this.txtMAC.Text += t_mac + Environment.NewLine;
            }

            foreach (string t_cpu in t_hardware.CPU)
            {
                this.txtCPU.Text += t_cpu + Environment.NewLine;
            }
        }
    }
}