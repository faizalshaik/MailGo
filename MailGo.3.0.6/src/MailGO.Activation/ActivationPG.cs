/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.23 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using OL = Microsoft.Office.Interop.Outlook;
using OC = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Globalization;
using System.Reflection;
// «[’Ç‰Á]
using Shell32;
using System.IO;
// ª[’Ç‰Á]
using OT = Microsoft.Office.Tools;

namespace DataDesign.MailGO.Activation
{
    public class ActivationPG : Model.MActivationPG
    {
        private OL.Application m_outlook = null;
        private Model.IEmail m_nocheck_email = null;
        private Model.IEmail m_nocheck_email1 = null;
        private OT.Ribbon.RibbonButton cmdMailGO = null;
        private OT.Ribbon.RibbonButton cmdOption = null;
        DateTime m_startedTime;

        private bool m_bActivateTemp;
        private bool m_statusTemp;

        public ActivationPG(Model.IMailGoPG v_mailgo, OL.Application v_outlook, OT.Ribbon.RibbonButton cmdBtn, OT.Ribbon.RibbonButton optBtn) : base(v_mailgo) 
        {
            this.m_mailgo.Track.Debug("ACTIVATION: Creating package ...");

            this.m_outlook = v_outlook;
            this.OnLoadOption();

            this.cmdMailGO = cmdBtn;
            this.cmdOption = optBtn;
            this.CreateToolbar();

            this.cmdMailGO.Enabled = true;
            this.cmdOption.Enabled = true;

            if (IsFirstUse())
            {
                string strLang = System.Threading.Thread.CurrentThread.CurrentCulture.EnglishName;
                if(strLang.IndexOf("Japanese") >=0 || strLang.IndexOf("Japan") >= 0)
                {
                    SetLanguage("ja-JP");
                }
                else
                    SetLanguage("en-US");
                SetFirstUse();
                //Should set default language.                
            }

            GetLanguage();
            if (m_language == "ja-JP")
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ja-JP");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("ja-JP");
            }
            else
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
            }

            m_startedTime = this.GetFirstUseTime();
            m_deadLine = m_startedTime.AddDays(14);
            //check license
            this.GetActivated();

            m_bActivatedReally = m_bActivated;

            if (m_bActivated ==false)
            {
                if (DateTime.Now < m_deadLine)
                    m_bActivated = true;
            }

            m_status = m_bActivated;
            refreshUIs();
        }

        private void SendEmail(Model.IEmail v_email)
        {
            OL._MailItem t_item =(OL._MailItem) this.m_outlook.CreateItem(OL.OlItemType.olMailItem);
            t_item.To = string.Join(";", v_email.TO.ToArray());
            t_item.CC = string.Join(";", v_email.CC.ToArray());
            t_item.Subject = v_email.Subject;
            t_item.Body = v_email.Body;
            //2009/6/6C³  OL.OlDefaultFolders.olFolderOutbox¨OL.OlDefaultFolders.olFolderSentMail
            t_item.SaveSentMessageFolder = this.m_outlook.Session.GetDefaultFolder(OL.OlDefaultFolders.olFolderSentMail);
            t_item.Send();
        }

        private string GetSender(OL.MailItem v_item)
        {
            string t_sender=null;
            try
            {
                if (v_item == null)
                {
                    t_sender = this.m_outlook.Session.CurrentUser.Address;
                }
                else
                {
                    if (this.m_outlook.Session.CurrentUser.AddressEntry.Type == "EX")
                    {
                        t_sender = GetSMTPAddressOfExchangeUser(v_item);
                    }
                    else
                    {
                        t_sender = this.m_outlook.Session.CurrentUser.Address;
                    }
                }

                this.m_mailgo.Track.Debug("ACTIVATION: Sender = " + (t_sender == null ? "NULL" : t_sender));
            }
            catch (Exception ex)
            {
                this.m_mailgo.Track.Error("GetSender ERROR: " + ex.Message);
                this.m_mailgo.Track.Error(ex);
                return t_sender;
            }
            return t_sender;
        }

        
        private bool OL2K7()
        {
            string[] fields = this.m_outlook.Version.Split(new char[] { '.' });
            int version = 0;
            int.TryParse(fields[0], out version);

            this.m_mailgo.Track.Debug("Version string: " + this.m_outlook.Version);
            this.m_mailgo.Track.Debug("Version number: " + version.ToString());
            
            return (version >= 12);
        }

        private void ClearFirstUse()
        {
            string t_subkey;
            string t_name;
            string t_boot = "1";

            this.GetBootKeys(out t_subkey, out t_name);
            this.SetValue(Registry.CurrentUser, t_subkey, t_name, t_boot);
        }

        private void SetActivated()
        {
            string t_subkey;
            string t_name;
            string t_activated = "1";

            this.GetActivatedKeys(out t_subkey, out t_name);
            this.SetValue(Registry.CurrentUser, t_subkey, t_name, t_activated);
        }
        private void SetDeActivated()
        {
            string t_subkey;
            string t_name;
            string t_activated = "0";

            this.GetActivatedKeys(out t_subkey, out t_name);
            this.SetValue(Registry.CurrentUser, t_subkey, t_name, t_activated);
        }

        private void GetLanguage()
        {
            string t_subkey = @"SOFTWARE\Kodai";
            string t_name = "LanguageMailGO";
            string t_language = "en-US";

            this.GetValue(out t_language, Registry.CurrentUser, t_subkey, t_name);
            if (t_language == "ja-JP")
                m_language = "ja-JP";
            else if(t_language == "en-US")
                m_language = "en-US";
            else 
            {
                string strLang = System.Threading.Thread.CurrentThread.CurrentCulture.EnglishName;
                if (strLang.IndexOf("Japanese") >= 0 || strLang.IndexOf("Japan") >= 0)
                    m_language = "ja-JP";
                else
                    m_language = "en-US";
                SetLanguage(m_language);
            }
        }
        private void SetLanguage(string name)
        {
            string t_subkey = @"SOFTWARE\Kodai";
            string t_name = "LanguageMailGO";
            this.SetValue(Registry.CurrentUser, t_subkey, t_name, name);
        }


        private void GetActivated()
        {
            string t_subkey;
            string t_name;
            string t_activated = "0";

            this.GetActivatedKeys(out t_subkey, out t_name);
            this.GetValue(out t_activated, Registry.CurrentUser, t_subkey, t_name);

            if (t_activated != null && t_activated == "1")
                m_bActivated = true;
            else
                m_bActivated = false;
        }

        private void SetFirstUseTime()
        {
            string t_startTime = DateTime.Now.ToString();
            this.SetValue(Registry.CurrentUser, @"SOFTWARE\Kodai", "StartTime", t_startTime);
        }

        private DateTime GetFirstUseTime()
        {
            string t_startTime;
            this.GetValue(out t_startTime, Registry.CurrentUser, @"SOFTWARE\Kodai", "StartTime");

            DateTime tTime = DateTime.Now;
            try
            {
                tTime = DateTime.Parse(t_startTime);
            }
            catch(Exception ee)
            {
                SetFirstUseTime();
            }
            return tTime;
        }

        private void SetFirstUse()
        {
            string t_subkey;
            string t_name;
            string t_boot = "0";

            this.GetBootKeys(out t_subkey, out t_name);
            this.SetValue(Registry.CurrentUser, t_subkey, t_name, t_boot);
        }

        private bool IsFirstUse()
        {
            string t_subkey;
            string t_name;
            string t_boot;

            this.GetBootKeys(out t_subkey, out t_name);
            this.GetValue(out t_boot, Registry.CurrentUser, t_subkey, t_name);

            return (t_boot == null || t_boot == "1");
        }

        private bool AskActivation()
        {
            DialogResult t_result;
            AskActivationWF t_form = new AskActivationWF();

            t_result = t_form.ShowDialog();

            return (t_result == DialogResult.Yes);
        }

        private bool AskDeActivation()
        {
            DialogResult t_result;
            AskDeActivationWF t_form = new AskDeActivationWF();

            t_result = t_form.ShowDialog();

            return (t_result == DialogResult.Yes);
        }


        private void SetOffMailGO()
        {
            this.m_mailgo.Track.Debug("ACTIVATION: Turn off MailGO ...");

            this.m_status = false;
            this.cmdMailGO.Label = ToolBarWF.Instance.lblMailGOTextOff.Text;
            this.cmdMailGO.ScreenTip = ToolBarWF.Instance.lblMailGOTipOff.Text;
            this.OnSaveOption();
        }

        private void SetOnMailGO()
        {
            this.m_mailgo.Track.Debug("ACTIVATION: Turn on MailGO ...");

            this.m_status = true;
            this.cmdMailGO.Label = ToolBarWF.Instance.lblMailGOTextOn.Text;
            this.cmdMailGO.ScreenTip = ToolBarWF.Instance.lblMailGOTipOn.Text;
            this.OnSaveOption();
        }

        private void UpdateMailGO()
        {
            this.m_mailgo.Track.Debug("ACTIVATION: Update MailGO Status ...");
            if (this.m_status)
            {
                this.cmdMailGO.Label = ToolBarWF.Instance.lblMailGOTextOn.Text;
                this.cmdMailGO.ScreenTip = ToolBarWF.Instance.lblMailGOTipOn.Text;
            }
            else
            {
                this.cmdMailGO.Label = ToolBarWF.Instance.lblMailGOTextOff.Text;
                this.cmdMailGO.ScreenTip = ToolBarWF.Instance.lblMailGOTipOff.Text;
            }
            this.m_mailgo.Track.Debug("ACTIVATION: END UpdateMailGO ");
        }

        private void CreateToolbar()
        {
            this.m_mailgo.Track.Debug("ACTIVATION: Creating toolbar ...");
            /*
            this.cmdMailGO.Label = ToolBarWF.Instance.lblMailGOTextOn.Text;
            this.cmdMailGO.ScreenTip = ToolBarWF.Instance.lblMailGOTipOn.Text;

            this.cmdOption.Label = ToolBarWF.Instance.lblOptionTextOn.Text;
            this.cmdOption.ScreenTip = ToolBarWF.Instance.lblOptionTipOn.Text;
            */


            /*
                        this.cmdMailGO.Click += new OC._CommandBarButtonEvents_ClickEventHandler(cmdMailGO_Click);
                        this.UpdateMailGO();

                        this.cmdOption = (OC.CommandBarButton)this.m_command_bar.Controls.Add(OC.MsoControlType.msoControlButton, 1, "", this.m_command_bar.Controls.Count + 1, false);

                        this.cmdOption.Caption = ToolBarWF.Instance.lblOptionTextOn.Text;

                        this.cmdOption.Style = OC.MsoButtonStyle.msoButtonCaption;

                        this.cmdOption.Click += new OC._CommandBarButtonEvents_ClickEventHandler(cmdOption_Click);            

                        //this.m_outlook.ItemSend += new OL.ApplicationEvents_10_ItemSendEventHandler(Outlook_ItemSend);
                        Outlook10EventHelper outlookHeper = new Outlook10EventHelper();
                        outlookHeper.SetupConnection(this.m_outlook, this);
            */
            this.m_mailgo.Track.Debug("ACTIVATION:// END Creating toolbar ");
        }

        private bool NoCheckEmail(Model.IEmail v_email)
        {
            this.m_mailgo.Track.Debug("ACTIVATION: Is no-check email?");

            if (this.m_nocheck_email == null) return false;
            
            Model.IEmail t_email = this.m_nocheck_email;
            this.m_nocheck_email = null;
            string t_target;
            string t_source;

            this.m_mailgo.Track.Debug("ACTIVATION: Check TO list");
            t_source = string.Join(";", v_email.TO.ToArray()).ToLower();
            t_target = string.Join(";", t_email.TO.ToArray()).ToLower();
            this.m_mailgo.Track.Debug("ACTIVATION: Source = " + t_source + " . Target = " + t_target);
            if (t_source != t_target) return false;

            this.m_mailgo.Track.Debug("ACTIVATION: Check CC list");
            t_source = string.Join(";", v_email.CC.ToArray()).ToLower();
            t_target = string.Join(";", t_email.CC.ToArray()).ToLower();
            this.m_mailgo.Track.Debug("ACTIVATION: Source = " + t_source + " . Target = " + t_target);
            if (t_source != t_target) return false;

            this.m_mailgo.Track.Debug("ACTIVATION: No-check email!");
            return true;
        }

        private bool NoCheckEmail1(Model.IEmail v_email)
        {
            this.m_mailgo.Track.Debug("ACTIVATION: Is no-check email?");

            if (this.m_nocheck_email1 == null) return false;

            Model.IEmail t_email = this.m_nocheck_email1;
            this.m_nocheck_email1 = null;
            string t_target;
            string t_source;

            this.m_mailgo.Track.Debug("ACTIVATION: Check TO list");
            t_source = string.Join(";", v_email.TO.ToArray()).ToLower();
            t_target = string.Join(";", t_email.TO.ToArray()).ToLower();
            this.m_mailgo.Track.Debug("ACTIVATION: Source = " + t_source + " . Target = " + t_target);
            if (t_source != t_target) return false;

            this.m_mailgo.Track.Debug("ACTIVATION: Check CC list");
            t_source = string.Join(";", v_email.CC.ToArray()).ToLower();
            t_target = string.Join(";", t_email.CC.ToArray()).ToLower();
            this.m_mailgo.Track.Debug("ACTIVATION: Source = " + t_source + " . Target = " + t_target);
            if (t_source != t_target) return false;

            this.m_mailgo.Track.Debug("ACTIVATION: No-check email!");
            return true;
        }


        public void Outlook_ItemSend(object v_item, ref bool v_cancel)
        {

            this.m_mailgo.Track.Debug("ACTIVATION: Outlook's email sent event handler ...");

            if (!this.m_mailgo.Activated) return;

            this.m_mailgo.Track.Debug("ACTIVATION: License is installed and status is on!");
            
            OL.MailItem t_item = v_item as OL.MailItem;
            Model.IEmail t_email = new Model.MEmail();

            t_email.Sender = this.GetSender(t_item);
            t_email.Body = t_item.Body;


            /* 2011/10/30 modified to handle null by msekine */
            if (t_email.Body == null)
            {
                t_email.Body = "\r\n";
            }

            t_email.Subject = t_item.Subject;

            //2011/10/11 ‚»‚à‚»‚àMailGo Status‚ªfalse‚Ìê‡‚Íƒ`ƒFƒbƒN‚µ‚È‚¢
            if (!this.m_status) return;

            foreach (OL.Recipient t_recipient in t_item.Recipients)
            {
                string address = string.Empty;
                if (t_recipient.AddressEntry.Type == "EX")
                    address = GetSMTPAddressOfExchangeUser(t_item);
                else
                    address = t_recipient.Address;

                if (t_recipient.Type == (int)OL.OlMailRecipientType.olTo)
                {
                    t_email.TO.Add(address);
                }

                if (t_recipient.Type == (int)OL.OlMailRecipientType.olCC)
                {
                    t_email.CC.Add(address);
                }
            }

            if (this.NoCheckEmail(t_email))
            {
                ConfirmRequestWF t_form = new ConfirmRequestWF();
                DialogResult t_result = t_form.ShowDialog();
                if (t_result != DialogResult.Yes)
                {
                    this.m_status = m_statusTemp;
                    this.m_bActivated = m_bActivateTemp;
                    v_cancel = true;
                }
                else
                {
                    this.m_bActivatedReally = true;
                    this.SetActivated();
                    UpdateMailGO();
                }
                return;
            }
            else if (this.NoCheckEmail1(t_email))
            {
                ConfirmRequestWF t_form = new ConfirmRequestWF();
                DialogResult t_result = t_form.ShowDialog();
                if (t_result != DialogResult.Yes)
                {
                    this.m_status = m_statusTemp;
                    this.m_bActivated = m_bActivateTemp;
                    v_cancel = true;
                }
                else
                {
                    this.m_bActivated = false;
                    this.m_bActivatedReally = false;
                    this.SetDeActivated();
                    if (DateTime.Now < m_deadLine)
                        m_bActivated = true;
                    m_status = m_bActivated;
                    UpdateMailGO();
                }
                return;
            }


            if (!this.m_status) return;

            this.OnBeforeSendEmail(t_email, out v_cancel);

            // «[’Ç‰Á]“Y•tƒtƒ@ƒCƒ‹ƒvƒƒpƒeƒBƒ`ƒFƒbƒN‹@”\
            // Šù‘¶ƒ`ƒFƒbƒN‚ÅƒLƒƒƒ“ƒZƒ‹‚µ‚Ä‚¢‚éê‡‚ÍAƒ[ƒ‹‘—M‚ðƒLƒƒƒ“ƒZƒ‹‚µ‚ÄI—¹‚·‚éB
            if (v_cancel == true)
            {
                return;
            }
            
            // 2009/9/18 ƒtƒ@ƒCƒ‹ƒvƒƒpƒeƒBƒIƒvƒVƒ‡ƒ“‚Ìƒ`ƒFƒbƒN
            if (this.m_option.CheckWord == false &&
                this.m_option.CheckExcel == false &&
                this.m_option.CheckPowerPoint == false &&
                this.m_option.CheckText == false ) return;

            // OS‚Ìƒo[ƒWƒ‡ƒ“‚ðŠm”F‚·‚éB
            String OS_Version = "";
            OperatingSystem osInfo = Environment.OSVersion;

            if (osInfo.Platform == PlatformID.Win32NT && osInfo.Version.Major == 5 && osInfo.Version.Minor == 1)
            {
                OS_Version = "Windows XP";
            }
            else if (osInfo.Platform == PlatformID.Win32NT && osInfo.Version.Major == 6 && osInfo.Version.Minor == 0)
            {
                OS_Version = "Windows Vista";
            }
            // 2010/9/13
            else if (osInfo.Platform == PlatformID.Win32NT && osInfo.Version.Major == 6 && osInfo.Version.Minor == 1)
            {
                OS_Version = "Windows 7";
            }
            else if(osInfo.Version.Major == 6 && osInfo.Version.Minor == 2)
                OS_Version = "Windows 8.0";
            else if (osInfo.Version.Major == 6 && osInfo.Version.Minor == 3)
                OS_Version = "Windows 8.1";
            else if (osInfo.Version.Major == 10 && osInfo.Version.Minor == 0)
                OS_Version = "Windows 10";

            //

            // “Y•tƒtƒ@ƒCƒ‹‚Ì—L–³‚ðŠm”F‚·‚éB
            int cnt = 0;
            foreach (OL.Attachment objAttachment in t_item.Attachments)
            {
                cnt++;
            }

            // ˆÈ‰º‚Ì‚¢‚Ã‚ê‚©‚ÉŠY“–‚·‚é‚Æ‚«‚É“Y•tƒtƒ@ƒCƒ‹ƒ`ƒFƒbƒN‚ðs‚¤B
            // EOS‚Ìƒo[ƒWƒ‡ƒ“‚ªWindows XP‚Å‚©‚Â“Y•tƒtƒ@ƒCƒ‹—L‚Ìê‡
            // EOS‚Ìƒo[ƒWƒ‡ƒ“‚ªWindows Vista‚Å‚©‚Â“Y•tƒtƒ@ƒCƒ‹—L‚Ìê‡
            if(cnt > 0 && (OS_Version == "Windows XP" || OS_Version == "Windows Vista" || OS_Version == "Windows 7"||
                OS_Version == "Windows 8.0" || OS_Version == "Windows 8.1" || OS_Version == "Windows 10"))
            {
                // ƒ[ƒ‹–{•¶‚Ìˆ¶æ‚ðŽæ“¾‚·‚éB
                string t_company = this.ReadCompany(t_email.Body);

                // Outlook‚Ìƒo[ƒWƒ‡ƒ“‚ðŽæ“¾‚·‚éB
                string oKeyName = @"Outlook.Application\CurVer";
                string oGetValueName = "";
                string outlook_version = "";
                try
                {
                    RegistryKey rKey = Registry.ClassesRoot.OpenSubKey(oKeyName);
                    outlook_version = (string)rKey.GetValue(oGetValueName);
                    rKey.Close();
                }
                catch (NullReferenceException)
                {
                    this.m_mailgo.Track.Debug("ƒŒƒWƒXƒgƒŠm" + oKeyName + "n‚Ìm" + oGetValueName + "n‚ª‚ ‚è‚Ü‚¹‚ñB");
                }

                // Outlook‚Ìƒo[ƒWƒ‡ƒ“‚É‘Î‰ž‚·‚éƒŒƒWƒXƒgƒŠƒL[‚ðÝ’è‚·‚éB
                string rKeyName = "";
                // Outlook 2002
                if (outlook_version == "Outlook.Application.10")
                {
                    rKeyName = @"Software\Microsoft\Office\10.0\Outlook\Security";
                }
                // Outlook 2003
                else if (outlook_version == "Outlook.Application.11") 
                {
                    rKeyName = @"Software\Microsoft\Office\11.0\Outlook\Security";
                }
                // Outlook 2007
                else if (outlook_version == "Outlook.Application.12")
                {
                    rKeyName = @"Software\Microsoft\Office\12.0\Outlook\Security";
                }
                // OutLook 2010  2010/09/13
                else if (outlook_version == "Outlook.Application.14")
                {
                    rKeyName = @"Software\Microsoft\Office\14.0\Outlook\Security";
                }
                else if (outlook_version == "Outlook.Application.16")
                {
                    rKeyName = @"Software\Microsoft\Office\16.0\Outlook\Security";
                }
                // ‚»‚Ì‘¼
                else
                {
                    // “Y•tƒtƒ@ƒCƒ‹ƒvƒƒpƒeƒBƒ`ƒFƒbƒN‹@”\‚ðI—¹‚·‚éB
                    return;
                }

                string rGetValueName = "OutlookSecureTempFolder";
                string location = "";

                try
                {
                    RegistryKey rKey = Registry.CurrentUser.OpenSubKey(rKeyName);
                    location = (string)rKey.GetValue(rGetValueName);
                    rKey.Close();
                }
                catch (NullReferenceException)
                {
                    this.m_mailgo.Track.Debug("ƒŒƒWƒXƒgƒŠm" + rKeyName + "n‚Ìm" + rGetValueName + "n‚ª‚ ‚è‚Ü‚¹‚ñB");
                }
                
                // “Y•tƒtƒ@ƒCƒ‹”‚¾‚¯ˆÈ‰º‚ðŒJ‚è•Ô‚·B
                foreach (OL.Attachment objAttachment in t_item.Attachments)
                {
                    // @Šg’£Žqƒ`ƒFƒbƒN@2010/9/18
                    String t_extention = "";
                    t_extention = Path.GetExtension(objAttachment.FileName);

                    if (t_extention == ".doc")
                    {
                        if (this.m_option.CheckWord == true)
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛB‰½‚à‚µ‚È‚¢‚ÅƒXƒ‹[‚·‚éB
                        }
                        else
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛŠOB
                            continue;
                        }
                    }
                    else if (t_extention == ".docx")
                    {
                        if (this.m_option.CheckWord == true)
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛB‰½‚à‚µ‚È‚¢‚ÅƒXƒ‹[‚·‚éB
                        }
                        else
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛŠOB
                            continue;
                        }
                    }
                    else if (t_extention == ".xls")
                    {
                        if (this.m_option.CheckExcel == true)
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛB‰½‚à‚µ‚È‚¢‚ÅƒXƒ‹[‚·‚éB
                        }
                        else
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛŠOB
                            continue;
                        }
                    }
                    else if (t_extention == ".xlsx")
                    {
                        if (this.m_option.CheckExcel == true)
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛB‰½‚à‚µ‚È‚¢‚ÅƒXƒ‹[‚·‚éB
                        }
                        else
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛŠOB
                            continue;
                        }
                    }
                    else if (t_extention == ".csv")
                    {
                        if (this.m_option.CheckExcel == true)
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛB‰½‚à‚µ‚È‚¢‚ÅƒXƒ‹[‚·‚éB
                        }
                        else
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛŠOB
                            continue;
                        }
                    }
                    else if (t_extention == ".ppt")
                    {
                        if (this.m_option.CheckPowerPoint == true)
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛB‰½‚à‚µ‚È‚¢‚ÅƒXƒ‹[‚·‚éB
                        }
                        else
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛŠOB
                            continue;
                        }
                    }
                    else if (t_extention == ".pptx")
                    {
                        if (this.m_option.CheckPowerPoint == true)
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛB‰½‚à‚µ‚È‚¢‚ÅƒXƒ‹[‚·‚éB
                        }
                        else
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛŠOB
                            continue;
                        }
                    }
                    else if (t_extention == ".txt")
                    {
                        if (this.m_option.CheckText == true)
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛB‰½‚à‚µ‚È‚¢‚ÅƒXƒ‹[‚·‚éB
                        }
                        else
                        {
                            // ƒ`ƒFƒbƒN‘ÎÛŠOB
                            continue;
                        }
                    }
                    else
                    {
                        // ƒ`ƒFƒbƒN‘ÎÛŠOB
                        continue;
                    }

                    // @“Y•tƒtƒ@ƒCƒ‹‚ÌƒvƒƒpƒeƒBuƒ^ƒCƒgƒ‹v‚Ì’l‚ðŽæ“¾‚·‚éB
                   
                    int TITLE_ID = 21;
                    // OS‚ªWindows XP‚Ìê‡
                    if (OS_Version == "Windows XP")
                    {
                        TITLE_ID = 10;
                    }
/*
                    // OS‚ªWindows Vista‚Ìê‡
                    else if (OS_Version == "Windows Vista")
                    {
                        TITLE_ID = 21;
                    }
                    // OS‚ªWindows 7‚Ìê‡ 2010/09/13
                    else if (OS_Version == "Windows 7")
                    {
                        TITLE_ID = 21;
                    }

                    else if (OS_Version == "Windows 8.0")
                    {
                        TITLE_ID = 22;
                    }
                    else if (OS_Version == "Windows 8.1")
                    {
                        TITLE_ID = 23;
                    }
                    else if (OS_Version == "Windows 10")
                    {
                        TITLE_ID = 31;
                    }
*/

                    ShellClass shell = new ShellClass();
                    Folder folder = shell.NameSpace(location);
                    FolderItem folderItem = folder.ParseName(objAttachment.FileName);

                    String t_title = "";
                    if (folderItem != null)
                    {
                        t_title = folder.GetDetailsOf(folderItem, TITLE_ID);
                        folderItem = null;
                    }
                    folder = null;
                    shell = null;
                    
                    // 
                    String l_title = "";
                    String l_company = "";

                    try
                    {
                        l_title = t_title.Trim().ToLower();
                    }
                    catch (NullReferenceException)
                    {
                        l_title = "";
                        t_title = "";
                    }

                    try
                    {
                        l_company = t_company.Trim().ToLower();
                    }
                    catch (NullReferenceException)
                    {
                        l_company = "";
                        t_company = "";
                    }
                    v_cancel = false;

                    /*

                    // ƒ[ƒ‹–{•¶‚Ìˆ¶æ‚Æuƒ^ƒCƒgƒ‹v‚Ì’l‚ð”äŠr‚·‚éB
                    if (l_company == l_title)
                    {
                        // ˆê’v‚·‚éê‡A‰½‚à‚µ‚È‚¢‚ÅƒXƒ‹[‚·‚éB

                    }
                    else
                    {
                        // ˆê’v‚µ‚È‚¢ê‡AŠm”Fƒ_ƒCƒAƒƒO‚ð•\Ž¦‚·‚éB                       
                        System.Windows.Forms.DialogResult t_result;
                        NotMatchWF t_form = new NotMatchWF(objAttachment.FileName, t_title, t_company);
                        t_result = t_form.ShowDialog();

                        // ‘—MŽÀsƒ{ƒ^ƒ“‰Ÿ‰ºŽžA‚»‚Ì‚Ü‚Ü‘±sB
                        if (t_result == System.Windows.Forms.DialogResult.OK)
                        {
                            v_cancel = false;
                        }

                        // ‘—M’†Ž~ƒ{ƒ^ƒ“‰Ÿ‰ºŽžA‘—M’†Ž~ˆ—‚ðŽÀŽ{B
                        if (t_result == System.Windows.Forms.DialogResult.Cancel)
                        {
                            v_cancel = true;
                            return;
                        }                       
                    }
                    */
                                      
                }

            }
            // ª[’Ç‰Á]“Y•tƒtƒ@ƒCƒ‹ƒvƒƒpƒeƒBƒ`ƒFƒbƒN‹@”\
            this.m_mailgo.Track.Debug("Cancel sending mail with v_cancel = " + v_cancel);
        }

        // «[’Ç‰Á]“Y•tƒtƒ@ƒCƒ‹ƒvƒƒpƒeƒBƒ`ƒFƒbƒN‹@”\
        private string ReadCompany(string v_body)
        {
            const string c_signals = " 　@\t\n\r";

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
        // ª[’Ç‰Á]“Y•tƒtƒ@ƒCƒ‹ƒvƒƒpƒeƒBƒ`ƒFƒbƒN‹@”\


        private string GetSMTPAddressOfExchangeUser(OL.MailItem mail)
        {
            return mail.SenderEmailAddress;
            //MessageBox.Show(mail.SenderEmailAddress);
            /*
                        const uint PR_SMTP_ADDRESS = 0x39FE001E;

                        MAPI.SessionClass objSession = new MAPI.SessionClass();
                        objSession.MAPIOBJECT = recipient.Application.Session.MAPIOBJECT;
                        MAPI.AddressEntry addEntry = (MAPI.AddressEntry)objSession.GetAddressEntry(recipient.EntryID);            
                        MAPI.Field field = (MAPI.Field)((MAPI.Fields)addEntry.Fields).get_Item(PR_SMTP_ADDRESS, null);

                        return field.Value.ToString();
            */
            string PR_SMTP_ADDRESS =
                @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            if (mail == null)
            {
                throw new ArgumentNullException();
            }
            if (mail.SenderEmailType == "EX")
            {
                OL.AddressEntry sender =
                    mail.Sender;
                if (sender != null)
                {
                    //Now we have an AddressEntry representing the Sender
                    if (sender.AddressEntryUserType ==OL.OlAddressEntryUserType.olExchangeUserAddressEntry
                        || sender.AddressEntryUserType ==OL.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
                    {
                        //Use the ExchangeUser object PrimarySMTPAddress
                        OL.ExchangeUser exchUser =sender.GetExchangeUser();
                        if (exchUser != null)
                        {
                            return exchUser.PrimarySmtpAddress;
                        }
                        else
                        {
                            return null;
                        }
                    }
                    else
                    {
                        return sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS) as string;
                    }
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return mail.SenderEmailAddress;
            }
        }


        public void cmdMailGO_Click()
        {
            if (m_bActivated == false)
                return;

            if (this.m_status)
            {
                this.SetOffMailGO();
            }
            else
            {
                this.SetOnMailGO();
            }
        }

        public void cmdOption_Click()
        {
            OptionWF opt = new OptionWF(m_mailgo);
            opt.ShowDialog();
        }

        private void GetOptionKeys(out string v_subkey, out string v_name)
        {
            v_subkey = @"SOFTWARE\Kodai";
            v_name = "OptionMailGO";
        }

        private void GetBootKeys(out string v_subkey, out string v_name)
        {
            v_subkey = @"SOFTWARE\Kodai";
            v_name = "BootMailGO";
        }

        private void GetActivatedKeys(out string v_subkey, out string v_name)
        {
            v_subkey = @"SOFTWARE\Kodai";
            v_name = "ActivateMailGO";
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

        //private void GetValue(out string v_value, string v_subkey, string v_name)
        //{
        //    RegistryKey t_regkey = Registry.CurrentUser.CreateSubKey(v_subkey);
        //    v_value = (string)t_regkey.GetValue(v_name);
        //    t_regkey.Close();
        //}

        //private void SetValue(string v_subkey, string v_name, string v_value)
        //{
        //    RegistryKey t_regkey = Registry.CurrentUser.CreateSubKey(v_subkey);
        //    t_regkey.SetValue(v_name, v_value);
        //    t_regkey.Close();
        //}

        protected override void OnActivate()
        {
            if (!AskActivation())
                return;
            //should send mail to 
            Model.IEmail t_email = new Model.MEmail();
            t_email.TO.Add("activation@kodaicorp.com");
            t_email.Sender = this.GetSender(null);
            t_email.Body = "MailGo activation 3.0.6";

            m_statusTemp = this.m_status;
            m_bActivateTemp = this.m_bActivated;
            this.m_bActivated = true;
            this.m_status = true;
            this.m_nocheck_email = t_email;
            this.SendEmail(t_email);
        }

        protected override void OnDeActivate()
        {

            if (!AskDeActivation())
                return;

            //should send mail to 
            Model.IEmail t_email = new Model.MEmail();
            t_email.TO.Add("activation@kodaicorp.com");
            t_email.Sender = this.GetSender(null);
            t_email.Body = "MailGo deactivation 3.0.6.0";

            m_statusTemp = this.m_status;
            m_bActivateTemp = this.m_bActivated;
            this.m_nocheck_email1 = t_email;
            this.SendEmail(t_email);
        }


        protected override void OnCleanUp()
        {
            this.m_mailgo.Track.Debug("ACTIVATION: Start cleaning up ...");

            if (this.m_outlook != null)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(this.m_outlook);
                }
                finally
                {
                    this.m_outlook = null;
                }
            }
        }

        protected override void OnChangeOption()
        {
        }

        protected override void OnLoadOption()
        {
            this.m_mailgo.Track.Debug("ACTIVATION: Loading option ...");

            string t_subkey;
            string t_name;
            string t_option;

            this.m_option.CheckAllAddress = true;
            this.m_option.CheckOnlyFirstAddress = false;
            this.m_option.CheckCCLine = true;
            this.m_option.CheckSameSuffix = true;
            this.m_status = false;
            // 2010/9/18
            this.m_option.CheckWord = false;
            this.m_option.CheckExcel = false;
            this.m_option.CheckPowerPoint = false;
            this.m_option.CheckText = false;
            //

            this.GetOptionKeys(out t_subkey, out t_name);
            this.GetValue(out t_option, Registry.CurrentUser, t_subkey, t_name);

            // 2009/5/27
            // [ORG]
            // if (t_option == null || t_option.Length < 5)
            
            if (t_option == null || t_option.Length < 9)
            {
                this.OnSaveOption();
            }
            else
            {
                this.m_option.CheckAllAddress = (t_option[0] == '1');
                this.m_option.CheckOnlyFirstAddress = (t_option[1] == '1');
                this.m_option.CheckCCLine = (t_option[2] == '1');
                this.m_option.CheckSameSuffix = (t_option[3] == '1');
                this.m_status = (t_option[4] == '1');
                // 2010/9/18
                this.m_option.CheckWord = (t_option[5] == '1');
                this.m_option.CheckExcel = (t_option[6] == '1');
                this.m_option.CheckPowerPoint = (t_option[7] == '1');
                this.m_option.CheckText = (t_option[8] == '1');
                //

            }
        }

        protected override void OnSaveOption()
        {
            this.m_mailgo.Track.Debug("ACTIVATION: Start saving option ...");

            string t_subkey;
            string t_name;
            string t_option = "";

            this.GetOptionKeys(out t_subkey, out t_name);
            t_option += (this.m_option.CheckAllAddress ? '1' : '0');
            t_option += (this.m_option.CheckOnlyFirstAddress ? '1' : '0');
            t_option += (this.m_option.CheckCCLine ? '1' : '0');
            t_option += (this.m_option.CheckSameSuffix ? '1' : '0');
            t_option += (this.m_status ? '1' : '0');
            // 2010/9/18
            t_option += (this.m_option.CheckWord ? '1' : '0');
            t_option += (this.m_option.CheckExcel ? '1' : '0');
            t_option += (this.m_option.CheckPowerPoint ? '1' : '0');
            t_option += (this.m_option.CheckText ? '1' : '0');
            //            
            this.SetValue(Registry.CurrentUser, t_subkey, t_name, t_option);
        }

        protected override void OnBeforeSendEmail(DataDesign.MailGO.Model.IEmail v_email, out bool v_cancel)
        {
            this.m_mailgo.Track.Debug("ACTIVATION: Start checking email before sent ...");
            this.m_mailgo.Address.CheckEmail(v_email, out v_cancel);            
        }

        protected override void OnChangeLanguage()
        {
            //write to registry
            SetLanguage(m_language);
            if (m_language == "ja-JP")
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ja-JP");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("ja-JP");
            }
            else
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
            }

            refreshUIs();
        }

        void refreshUIs()
        {
            ToolBarWF.sm_instance = null;
            if (m_status == false)
            {
                this.cmdMailGO.Label = ToolBarWF.Instance.lblMailGOTextOff.Text;
                this.cmdMailGO.ScreenTip = ToolBarWF.Instance.lblMailGOTipOff.Text;
            }
            else
            {
                this.cmdMailGO.Label = ToolBarWF.Instance.lblMailGOTextOn.Text;
                this.cmdMailGO.ScreenTip = ToolBarWF.Instance.lblMailGOTipOn.Text;
            }
            this.cmdOption.Label = ToolBarWF.Instance.lblOptionTextOn.Text;
            this.cmdOption.ScreenTip = ToolBarWF.Instance.lblOptionTextOn.Text;
        }


    }
}
