/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.24 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.IO;
using System.Reflection;
using System.Xml;
using OL = Microsoft.Office.Interop.Outlook;
using OT = Microsoft.Office.Tools;
using System.Windows.Forms;

namespace DataDesign.MailGO.Unity
{
    internal class MailGoPG : Model.MMailGoPG
    {
        private OL.Application m_outlook;
        private string m_config_file;

        public OT.Ribbon.RibbonButton cmdMailGO;
        public OT.Ribbon.RibbonButton cmdOption;

        public MailGoPG(OL.Application v_outlook)
            : base()
        {
            this.m_outlook = v_outlook;
            Track.Configure(out this.m_config_file);
        }

        protected override void OnCreateTrack()
        {
            this.m_track = new Track();            
        }

        protected override void OnCreateActivation()
        {
            this.m_activation = Activation.Share.CreatePG(this, this.m_outlook, cmdMailGO, cmdOption);
            ((Model.IMailGoPG)this).Track.Debug("OnCreateActivation: END Activation.");
        }

        protected override bool OnCheckActivated()
        {
            return m_activation.Activated;
        }

        protected override void OnCreateAddress()
        {            
            ((Model.IMailGoPG)this).Track.Debug("OnCreateAddress() Begin.");
            this.m_address = Address.Share.CreatePG(this);

            string dirName = Path.Combine(Environment.GetFolderPath(
                Environment.SpecialFolder.ApplicationData), "MailGo\\");
            Directory.CreateDirectory(dirName);
            string fileName = Path.Combine(dirName, "UserDefined.csv");
            //MessageBox.Show(fileName);
            this.m_address.UserDefinedFile = fileName;

            /*
                        XmlDocument t_doc;
                        XmlNode t_node;
                        string t_path;

                        t_path = Path.GetDirectoryName(this.m_config_file);
                        t_doc = new XmlDocument();
                        t_doc.Load(this.m_config_file);

                        //t_node = t_doc.SelectSingleNode("/configuration/appSettings/add[@key=\"ShipmentList\"]/@value");
                        //this.m_address.ShipmentFile = Path.Combine(t_path, t_node.Value.ToString());
                        t_node = t_doc.SelectSingleNode("/configuration/appSettings/add[@key=\"UserDefinedList\"]/@value");
                        if (t_node != null && t_node.Value != null)
                        {
                            this.m_address.UserDefinedFile = Path.Combine(t_path, t_node.Value.ToString());
                        }
                        else
                        {
                            this.m_address.UserDefinedFile = String.Format("{0}\\UserDefined.csv", t_path);
                        }
                        ((Model.IMailGoPG)this).Track.Debug("OnCreateAddress() End.");
            */
        }

        protected override void OnStartup()
        {
            ((Model.IMailGoPG)this).Track.Debug("Startup MailGO ...");
            bool t_status = ((Model.IMailGoPG)this).Activation.Status;
        }

        protected override void OnShutdown()
        {
            ((Model.IMailGoPG)this).Track.Debug("Shutdown MailGO ...");
            this.m_activation.CleanUp();
        }
    }
}
