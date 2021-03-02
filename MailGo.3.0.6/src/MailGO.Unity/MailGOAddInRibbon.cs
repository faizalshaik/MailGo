using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Diagnostics;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace DataDesign.MailGO.Unity
{
   public partial class MailGOAddInRibbon
   {
      private Model.IMailGoPG m_mailgo;
      private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
      {
        log4net.Config.XmlConfigurator.Configure();
        m_mailgo = new MailGoPG(Globals.ThisAddIn.Application);
        MailGoPG mailGO = (MailGoPG)m_mailgo;

        mailGO.cmdMailGO = this.cmdMailGO;
        mailGO.cmdOption = this.cmdOption;
        m_mailgo.Startup();
        Globals.ThisAddIn.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(app_ItemSend);

        try
        {
            //this.group1.Label = String.Format(
            //this.group1.Label, "IT");
         }
         catch (System.Exception ex)
         {
            MessageBox.Show(ex.Message);
         }
      }

        private void cmdMailGO_Click(object sender, RibbonControlEventArgs e)
        {
            Activation.Share.onCmdMailGoClick();
        }

        private void cmdOption_Click(object sender, RibbonControlEventArgs e)
        {
            Activation.Share.onOptionClick();
        }
        void app_ItemSend(object Item, ref bool Cancel)
        {
            Outlook.MailItem item = Item as Outlook.MailItem;
            if (item != null)
            {
                Activation.Share.onBeforeSendMail(item , out Cancel);
            }
        }
    }
}
