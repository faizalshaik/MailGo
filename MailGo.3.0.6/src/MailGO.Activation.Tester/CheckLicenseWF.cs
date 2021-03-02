using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DataDesign.MailGO.Activation.Tester
{
    public partial class CheckLicenseWF : Form
    {
        private static bool sm_installed;
        private static bool sm_evaluted;
        private static bool sm_activated;
        private static bool sm_expired;
        private static DateTime? sm_expired_date;
        private static string sm_license_id;

        public CheckLicenseWF()
        {
            InitializeComponent();
        }

        public static void AskOutput(out bool v_installed, out bool v_evaluted, out bool v_activated, out bool v_expired, out DateTime? v_expired_date, out string v_license_id)
        {
            CheckLicenseWF t_form = new CheckLicenseWF();
            t_form.ShowDialog();

            v_installed = sm_installed;
            v_evaluted = sm_evaluted;
            v_activated = sm_activated;
            v_expired = sm_expired;
            v_expired_date = sm_expired_date;
            v_license_id = sm_license_id;
        }

        private void CheckLicenseWF_FormClosing(object sender, FormClosingEventArgs e)
        {
            sm_installed = this.chkInstalled.Checked;
            sm_evaluted = this.chkEvaluated.Checked;
            sm_activated = this.chkActivated.Checked;
            sm_expired = this.chkExpired.Checked;
            if (this.dtpExpiredDate.Checked)
            {
                sm_expired_date = this.dtpExpiredDate.Value;
            }
            else
            {
                sm_expired_date = null;
            }
            sm_license_id = this.txtLicenseID.Text;
            if (sm_license_id.Length == 0)
            {
                sm_license_id = null;
            }
        }
    }
}