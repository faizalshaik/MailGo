/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.23 Tran Dinh Thoai
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

namespace DataDesign.MailGO.Activation.Tester
{
    internal class MailGoPG : Model.MMailGoPG
    {
        private const string c_lic_activation = "awaoghoeiaghdkahgoekghanbd&322234dghoa&*dhgow*eygkah&dkag";

        private OL.Application m_outlook;
        public OT.Ribbon.RibbonButton cmdMailGO;
        public OT.Ribbon.RibbonButton cmdOption;


        public MailGoPG(OL.Application v_outlook)
            : base()
        {
            this.m_outlook = v_outlook;
            Track.Configure();
        }

        protected override void OnCreateTrack()
        {
            this.m_track = new Track();
        }

        protected override void OnCreateActivation()
        {
            this.m_activation = Activation.Share.CreatePG(this, this.m_outlook, cmdMailGO, cmdOption);
        }


        protected override void OnStartup()
        {
            ((Model.IMailGoPG)this).Track.Debug("Startup ...");
            bool t_status = ((Model.IMailGoPG)this).Activation.Status;
        }

        protected override void OnShutdown()
        {
            ((Model.IMailGoPG)this).Track.Debug("Shutdown ...");
        }
    }
}
