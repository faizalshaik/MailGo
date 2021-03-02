/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.21 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace DataDesign.MailGO.Address
{
    internal class Checker
    {
        private Model.IAddressPG m_address;
        

        public Checker(Model.IAddressPG v_address)
        {
            this.m_address = v_address;
        }

        public void Run(Model.IEmail v_email, out bool v_cancel)
        {
            v_cancel = true;            
            bool t_cancel;
            bool t_ignore;
            string t_email;
            string t_suffix;
            List<string> t_checklist = new List<string>();
            string t_sender_suffix = this.ReadSuffix(v_email.Sender).Trim().ToLower();
            string t_company = this.ReadCompany(v_email.Body);
                        
            if (v_email.Body.Trim() == string.Empty)
            {
                string message = DataDesign.MailGO.Utils.ResourceUtil.Instance.GetString("MSG_BODY_EMPTY");
                string caption = DataDesign.MailGO.Utils.ResourceUtil.Instance.GetString("PRODUCT_NAME");
                // 2009/5/30
                // ƒƒbƒZ[ƒW‚ðƒfƒXƒNƒgƒbƒv‚É‹­§•\Ž¦‚·‚é‚½‚ßAMessageBoxOptions.DefaultDesktopOnly‚ð’Ç‰ÁB
                DialogResult result = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly);
                if (result == DialogResult.No)
                {
                    v_cancel = true;
                    return;
                }
            }
            else
            {
                if (t_company.Length == 0)
                {
                    v_cancel = true;
                    return;
                }
            }

            if (this.m_address.MailGo.Activation.Option.CheckAllAddress)
            {
                t_checklist.AddRange(v_email.TO.ToArray());
            }
            else if (v_email.TO.Count > 0)
            {
                t_checklist.Add(v_email.TO[0]);
            }
            if (this.m_address.MailGo.Activation.Option.CheckCCLine)
            {
                t_checklist.AddRange(v_email.CC.ToArray());
            }

            for (int t_idx = t_checklist.Count - 1; t_idx >= 0; t_idx--)
            {
                t_email = t_checklist[t_idx];
                if (t_email.Trim().Length == 0)
                {
                    t_checklist.RemoveAt(t_idx);
                }
                else if (!this.m_address.MailGo.Activation.Option.CheckSameSuffix)
                {
                    t_suffix = this.ReadSuffix(t_email).Trim().ToLower();
                    if (t_suffix == t_sender_suffix)
                    {
                        t_checklist.RemoveAt(t_idx);
                        v_cancel = false;                        
                    }
                }
            }

            //MessageBox.Show("addr count " + t_checklist.Count.ToString());
            for (int t_idx = 0; t_idx < t_checklist.Count; t_idx++)
            {
                t_email = t_checklist[t_idx];
                t_suffix = this.ReadSuffix(t_email);

                //MessageBox.Show(t_idx.ToString()+ " th addr");
                if (this.Check(t_company, t_suffix))
                {
                    //MessageBox.Show(t_idx.ToString() + " th addr cheked!");
                    v_cancel = false;
                    //return;
                    continue;
                }

                this.Confirm(t_company, t_email, t_suffix, out t_cancel, out t_ignore);

                if (t_cancel)
                {
                    //MessageBox.Show("t-cancel");
                    v_cancel = true;
                    return;
                }
                else if (t_ignore)
                {
                    //MessageBox.Show("t-ignore");
                    if (t_idx == t_checklist.Count - 1)
                    {
                        v_cancel = false;
                        return;
                    }
                }
                else
                {
                    //MessageBox.Show("t-addlist");
                    v_cancel = false;
                    //return;
                    continue;
                }
            }
        }

        private void Confirm(string v_company, string v_email, string v_suffix, out bool v_cancel, out bool v_ignore)
        {
            v_cancel = false;
            v_ignore = false;

            System.Windows.Forms.DialogResult t_result;
            NoMatchWF t_form = new NoMatchWF(v_company, v_email);
            t_result = t_form.ShowDialog();

            if (t_result == System.Windows.Forms.DialogResult.Cancel)
            {
                v_cancel = true;
            }
            else if (t_result == System.Windows.Forms.DialogResult.Ignore)
            {
                v_ignore = true;
            }
            else
            {
                this.m_address.UserDefinedList.Add(new Model.MAddress(v_company, v_suffix));

                if (m_address.UserDefinedFile != null && m_address.UserDefinedFile != "")
                {
                    CSV t_csv = new CSV(this.m_address.UserDefinedFile);
                    t_csv.Write(this.m_address.UserDefinedList);
                }
            }
        }

        private bool Check(string v_company, string v_suffix)
        {
            //if (this.Check(v_company, v_suffix, this.m_address.ShipmentList)) return true;
            if (this.Check(v_company, v_suffix, this.m_address.UserDefinedList)) return true;
            return false;
        }

        private bool Check(string v_company, string v_suffix, List<Model.IAddress> v_address_list)
        {
            string t_company;
            string t_suffix;

            v_company = v_company.Trim().ToLower();
            v_suffix = v_suffix.Trim().ToLower();

            foreach (Model.IAddress t_address in v_address_list)
            {
                t_company = t_address.Company.Trim().ToLower();
                t_suffix = t_address.Suffix.Trim().ToLower();
                if (t_company == v_company && t_suffix == v_suffix) return true;
            }

            return false;
        }

        private string ReadSuffix(string v_email)
        {
            string t_suffix = v_email.Trim().ToLower();
            t_suffix = t_suffix.Substring(t_suffix.IndexOf('@') + 1);
            return t_suffix;
        }

        private string ReadCompany(string v_body)
        {
            const string c_signals = " 　@\t\n\r";
            
            /* 2011/10/30 modified to handle null by msekine */
            if (v_body == null)
            {
                v_body = "\r\n";
            }
            string msgBody = v_body.Trim();

            string t_company = "";
            if (msgBody == string.Empty)
                return t_company;            

            char t_char;            
            for (int t_idx = 0; t_idx < msgBody.Length; t_idx++)
            {
                t_char = msgBody[t_idx];
                if (c_signals.IndexOf(t_char) >= 0) break;
                t_company += t_char;
            }

            return t_company;
        }
    }
}
