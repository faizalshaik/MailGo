/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.23 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using OL = Microsoft.Office.Interop.Outlook;
using OT = Microsoft.Office.Tools;

namespace DataDesign.MailGO.Activation
{
    public class Share
    {
        private static ActivationPG g_ActivationPG = null;
        public static Model.IActivationPG CreatePG(Model.IMailGoPG v_mailgo, OL.Application v_outlook, OT.Ribbon.RibbonButton cmdBtn, OT.Ribbon.RibbonButton optBtn)
        {
            g_ActivationPG = new ActivationPG(v_mailgo, v_outlook, cmdBtn, optBtn);
            return g_ActivationPG;
        }

        public static void onCmdMailGoClick()
        {
            g_ActivationPG.cmdMailGO_Click();
        }
        public static void onOptionClick()
        {
            g_ActivationPG.cmdOption_Click();
        }

        public static void onBeforeSendMail(OL.MailItem v_email, out bool v_cancel)
        {
            bool bCancel = false;
            g_ActivationPG.Outlook_ItemSend(v_email, ref bCancel);
            v_cancel = bCancel;
        }
    }
}
