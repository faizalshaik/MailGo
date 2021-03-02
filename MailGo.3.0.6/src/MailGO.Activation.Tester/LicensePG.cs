/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.23 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;

namespace DataDesign.MailGO.Activation.Tester
{
    internal class LicensePG : Model.MLicensePG
    {
        public LicensePG(Model.IMailGoPG v_mailgo) : base(v_mailgo) { }

        protected override void OnCheckLicense(out bool v_installed, out bool v_evaluted, out bool v_activated, out bool v_expired, out DateTime? v_expired_date, out string v_license_id)
        {
            CheckLicenseWF.AskOutput(out v_installed, out v_evaluted, out v_activated, out v_expired, out v_expired_date, out v_license_id);
        }
    }
}
