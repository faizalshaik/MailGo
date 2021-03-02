/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.21 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

namespace DataDesign.MailGO.License
{
    internal class LicensePG : Model.MLicensePG
    {
        public LicensePG(Model.IMailGoPG v_mailgo) : base(v_mailgo) { }

        protected override DataDesign.MailGO.Model.IHardware OnReadHardware()
        {
            this.m_mailgo.Track.Debug("LICENSE: Reading hardware [started]");

            Model.IHardware t_hardware = new Model.MHardware();
            
            t_hardware.MAC.AddRange(Hardware.GetMACAddress().ToArray());
            t_hardware.CPU.AddRange(Hardware.GetCPUId().ToArray());

            this.m_mailgo.Track.Debug("LICENSE: Reading hardware [finished]");

            return t_hardware;
        }

        protected override void OnGenerateLicenseID()
        {
            this.m_mailgo.Track.Debug("LICENSE: Display generate license form ...");
            LicenseWF t_form = new LicenseWF(this.m_mailgo.License);
            t_form.ShowDialog();
        }

        protected override string OnGenerateLicenseID(bool v_evaluation, int v_activation_period)
        {
            this.m_mailgo.Track.Debug("LICENSE: Generating license id ...");

            char t_version_type;
            string t_activate_time;
            string t_guid = System.Guid.NewGuid().ToString();
            string t_license;

            this.GetVersionType(out t_version_type, v_evaluation);
            this.GetActivateTime(out t_activate_time, v_activation_period);
            this.GetLicenseId(out t_license, t_version_type, t_activate_time, t_guid);

            this.m_mailgo.Track.Debug("LICENSE: Raw license id string is created!");

            return this.EncryptLicenseId(t_license);
        }

        protected override void OnGenerateActivationID()
        {
            this.m_mailgo.Track.Debug("LICENSE: Display generate activation id form ...");
            ActivationWF t_form = new ActivationWF(this.m_mailgo.License);
            t_form.ShowDialog();
        }

        protected override string OnGenerateActivationID(string v_license_id, DataDesign.MailGO.Model.IHardware v_hardware)
        {
            this.m_mailgo.Track.Debug("LICENSE: Generating activation id ...");

            string t_license_id = this.DecryptLicenseId(v_license_id);
            string t_target = Hardware.Export(v_hardware) + "|" + t_license_id;

            this.m_mailgo.Track.Debug("LICENSE: Raw activation id string is created!");

            return this.EncryptActivationId(t_target);
        }

        protected override void OnRequestActivationID(out string v_license_id, out DataDesign.MailGO.Model.IHardware v_hardware, string v_request)
        {
            this.m_mailgo.Track.Debug("LICENSE: Request activation id [started]");

            int t_pos = v_request.IndexOf("|");
            v_license_id = v_request.Substring(t_pos + 1);
            v_hardware = Hardware.Import(v_request.Substring(0, t_pos));

            this.m_mailgo.Track.Debug("LICENSE: Request activation id [finished]");
        }

        protected override void OnRequestActivationID(out string v_request, string v_license_id, DataDesign.MailGO.Model.IHardware v_hardware)
        {
            this.m_mailgo.Track.Debug("LICENSE: Request activation id [started]");

            v_request = Hardware.Export(v_hardware) + "|" + v_license_id;

            this.m_mailgo.Track.Debug("LICENSE: Request activation id [finished]");
        }

        protected override void OnInstallLicense(string v_license_id)
        {
            this.m_mailgo.Track.Debug("LICENSE: Installing license [started]");

            bool t_installed;
            bool t_evaluated;
            bool t_activated;
            bool t_expired;
            DateTime? t_expired_date;
            string t_license_id;
            char t_version_type;
            string t_activate_period;
            string t_guid;
            bool t_new_evaluated;

            this.m_mailgo.License.CheckLicense(out t_installed);
            this.GetLicenseId(out t_version_type, out t_activate_period, out t_guid, this.DecryptLicenseId(v_license_id));
            this.GetVersionType(out t_new_evaluated, t_version_type);

            if (t_installed)
            {
                this.m_mailgo.Track.Debug("LICENSE: License is installed!");

                if (t_evaluated)
                {
                    this.m_mailgo.Track.Debug("LICENSE: Evaluation version!");
                    if (t_new_evaluated) return;
                }
                else
                {
                    this.m_mailgo.Track.Debug("LICENSE: Standard version!");
                    if (t_license_id == v_license_id) return;
                }
            }

            string v_subkey;
            string v_name;
            string t_usage_info;

            this.GetUsageInfoKeys(out v_subkey, out v_name);
            this.GetUsageInfo(out t_usage_info, DateTime.Now, DateTime.Now, null, ((Model.ILicensePG)this).ReadHardware(), null, this.DecryptLicenseId(v_license_id));
            this.SetValue(Registry.CurrentUser, v_subkey, v_name, this.EncryptUsageInfo(t_usage_info));
            this.SetFirstUse();

            this.m_mailgo.Track.Debug("LICENSE: Installing license [finished]");
        }
        protected override void OnActivateLicense(string v_activation_id)
        {
            string[] t_fields;

            string v_subkey;
            string v_name;
            string t_usage_info;
            DateTime? t_install_time = null;
            DateTime? t_update_time = null;
            DateTime? t_activate_time = null;
            Model.IHardware t_install_hardware = null;
            Model.IHardware t_activate_hardware = null;
            string t_license_id;
            string t_activate_license_id;

            this.m_mailgo.Track.Debug("Start activating license");

            this.GetUsageInfoKeys(out v_subkey, out v_name);
            this.GetValue(out t_usage_info, Registry.LocalMachine, v_subkey, v_name);

            t_usage_info = this.DecryptUsageInfo(t_usage_info);
            this.GetUsageInfo(out t_install_time, out t_update_time, out t_activate_time, out t_install_hardware, out t_activate_hardware, out t_license_id, t_usage_info);

            try
            {
                string t_raw_activation_id = this.DecryptActivationId(v_activation_id);
                t_fields = t_raw_activation_id.Split(new char[] { '|' });
                t_activate_hardware = Hardware.Import(t_fields[0]);
                t_activate_license_id = t_fields[1];
            }
            catch (Exception e)
            {
                this.m_mailgo.Track.Error(e);
                throw new Model.XBadActivationID();
            }
            if (!Hardware.IsSame(t_install_hardware, t_activate_hardware))
            {
                throw new Model.XBadActivationID();
            }
            if (t_activate_license_id != t_license_id)
            {
                throw new Model.XBadActivationID();
            }
            t_update_time = DateTime.Now;
            t_activate_time = DateTime.Now;
            this.GetUsageInfo(out t_usage_info, t_install_time, t_update_time, t_activate_time, t_install_hardware, t_activate_hardware, t_license_id);
            this.SetValue(Registry.CurrentUser, v_subkey, v_name, this.EncryptUsageInfo(t_usage_info));

            this.m_mailgo.Track.Debug("Finish activating license");
        }

        protected override void OnActivateLicense()
        {
            this.m_mailgo.Track.Debug("LICENSE: Display activate license form ...");

            ActivateLicenseWF t_form = new ActivateLicenseWF(this.m_mailgo);
            t_form.ShowDialog();
        }

        protected override void OnInstallLicense()
        {
            this.m_mailgo.Track.Debug("LICENSE: Display install license form ...");

            InstallLicenseWF t_form = new InstallLicenseWF(this.m_mailgo);
            t_form.ShowDialog();
        }

        protected override bool OnCheckCDO()
        {
            bool t_result = true;
            bool t_ol2k7 = IsOL2K7();

            if (!this.CDOInstalled(t_ol2k7))
            {
                t_result = false;
                CDORequiredWF.Show(t_ol2k7);
            }

            return t_result;
        }

        private bool IsOL2K7()
        {
            const string c_subkey = @"SOFTWARE\Microsoft\Office\12.0\Outlook\InstallRoot";
            const string c_name = @"Path";

            string t_value = null;

            GetValue(out t_value, Registry.LocalMachine, c_subkey, c_name);

            return (t_value != null && System.IO.Directory.Exists(t_value));
        }

        private bool CDOInstalled(bool v_ol2k7)
        {
            const string c_subkey = @"MAPI.Session";
            const string c_name = null;

            string t_value = null;

            GetValue(out t_value, Registry.ClassesRoot, c_subkey, c_name);

            return (t_value != null);
        }

        private void GetUsageInfoKeys(out string v_subkey, out string v_name)
        {
            v_subkey = @"SOFTWARE\Kodai";
            v_name = "UseMailGO";
        }

        private void GetValue(out string v_value, RegistryKey v_key, string v_subkey, string v_name)
        {
            RegistryKey t_regkey = v_key.OpenSubKey(v_subkey);
            if (t_regkey == null)
            {
                v_value = null;
            } 
            else 
            {
                v_value = (string)t_regkey.GetValue(v_name);
                t_regkey.Close();
            }
        }

        private void SetValue(RegistryKey v_key, string v_subkey, string v_name, string v_value)
        {
            RegistryKey t_regkey = v_key.CreateSubKey(v_subkey);
            t_regkey.SetValue(v_name, v_value);
            t_regkey.Close();
        }

        private string DecryptActivationId(string v_source)
        {
            const string c_iv = "a&#@";
            const string c_key = "pwgow&d%(*a!e^+h";
            return Crypto.Decrypt(v_source, c_key, c_iv);
        }

        private string EncryptActivationId(string v_source)
        {
            const string c_iv = "a&#@";
            const string c_key = "pwgow&d%(*a!e^+h";
            return Crypto.Encrypt(v_source, c_key, c_iv);
        }

        private string DecryptLicenseId(string v_source)
        {
            const string c_iv = "$q!w";
            const string c_key = "ag^dow&sk#doa%dh";
            return Crypto.Decrypt(v_source, c_key, c_iv);
        }

        private string EncryptLicenseId(string v_source)
        {
            const string c_iv = "$q!w";
            const string c_key = "ag^dow&sk#doa%dh";
            return Crypto.Encrypt(v_source, c_key, c_iv);
        }

        private string DecryptUsageInfo(string v_source)
        {
            const string c_iv = "#a&%";
            const string c_key = "!@thgowh&*d*^)$d";
            return Crypto.Decrypt(v_source, c_key, c_iv);
        }

        private string EncryptUsageInfo(string v_source)
        {
            const string c_iv = "#a&%";
            const string c_key = "!@thgowh&*d*^)$d";
            return Crypto.Encrypt(v_source, c_key, c_iv);
        }

        private void GetUsageInfo(out string v_target, DateTime? v_install_time, DateTime? v_update_time, DateTime? v_activate_time, Model.IHardware v_install_hardware, Model.IHardware v_activate_hardware, string v_raw_license_id)
        {
            const string c_format = "fffssmmHHddMMyyyy";

            v_target = (v_install_time.HasValue ? v_install_time.Value.ToString(c_format) : "");
            v_target += "|" + (v_update_time.HasValue ? v_update_time.Value.ToString(c_format) : "");
            v_target += "|" + (v_activate_time.HasValue ? v_activate_time.Value.ToString(c_format) : "");
            v_target += "|" + (v_install_hardware != null ? Hardware.Export(v_install_hardware) : "");
            v_target += "|" + (v_activate_hardware != null ? Hardware.Export(v_activate_hardware) : "");
            v_target += "|" + v_raw_license_id;
        }

        private void GetUsageInfo(out DateTime? v_install_time, out DateTime? v_update_time, out DateTime? v_activate_time, out Model.IHardware v_install_hardware, out Model.IHardware v_activate_hardware, out string v_raw_license_id, string v_source)
        {
            v_install_time = null;
            v_update_time = null;
            v_activate_time = null;
            v_install_hardware = null;
            v_activate_hardware = null;
            v_raw_license_id = null;

            string[] t_fields = v_source.Split(new char[] { '|' });

            if (t_fields.Length < 6) return;

            v_install_time = ParseDateTime(t_fields[0]);
            v_update_time = ParseDateTime(t_fields[1]);
            v_activate_time = ParseDateTime(t_fields[2]);
            v_install_hardware = ParseHardware(t_fields[3]);
            v_activate_hardware = ParseHardware(t_fields[4]);

            v_raw_license_id = "";
            for (int t_idx = 5; t_idx < t_fields.Length; t_idx++)
            {
                v_raw_license_id += "|" + t_fields[t_idx];
            }
            if (v_raw_license_id.Length > 0)
            {
                v_raw_license_id = v_raw_license_id.Substring(1);
            }
        }

        private Model.IHardware ParseHardware(string v_source)
        {
            Model.IHardware t_target = null;
            try
            {
                t_target = Hardware.Import(v_source);
            }
            catch { }

            return t_target;
        }

        private DateTime? ParseDateTime(string v_source)
        {
            const string c_format = "fffssmmHHddMMyyyy";

            DateTime? t_target = null;
            try
            {
                t_target = DateTime.ParseExact(v_source, c_format, null);
            }
            catch { }

            return t_target;
        }

        private void GetLicenseId(out string v_target, char v_version_type, string v_activate_time, string v_guid)
        {
            const string c_format = "fffssmmHHddMMyyyy";
            v_target = DateTime.Now.ToString(c_format) + v_version_type + v_activate_time + v_guid;
        }

        private void GetLicenseId(out char v_version_type, out string v_activate_time, out string v_guid, string v_source)
        {
            v_source = v_source.Substring(17);
            v_version_type = v_source[0];
            v_activate_time = v_source.Substring(1, 10);
            v_guid = v_source.Substring(11);
        }

        private void GetActivateTime(out string v_target, int v_source)
        {
            v_target = v_source.ToString().PadLeft(10, ' ');
        }

        private void GetActivateTime(out int v_target, string v_source)
        {
            v_target = int.Parse(v_source);
        }

        private void GetVersionType(out char v_type, bool v_evaluation)
        {
            v_type = (v_evaluation ? 'E' : 'S');
        }

        private void GetVersionType(out bool v_evaluation, char v_type)
        {
            v_evaluation = (v_type == 'E');
        }

        private void SetFirstUse()
        {
            string t_subkey;
            string t_name;
            string t_boot = "0";

            this.GetBootKeys(out t_subkey, out t_name);
            this.SetValue(Registry.CurrentUser, t_subkey, t_name, t_boot);
        }

        private void GetBootKeys(out string v_subkey, out string v_name)
        {
            v_subkey = @"SOFTWARE\Kodai";
            v_name = "BootMailGO";
        }
    }
}
