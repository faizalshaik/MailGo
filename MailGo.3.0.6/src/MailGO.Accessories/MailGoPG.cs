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

namespace DataDesign.MailGO.Accessories
{
    internal class MailGoPG : Model.MMailGoPG
    {
        private const string c_lic_license = "aoekghoakwhgodkahgoekhgaowighke1234dfhe@343tdogh&68*964546dhgoqdkgha";

        protected override void OnCreateTrack()
        {
            this.m_track = new Track();
        }

        protected override void OnStartup()
        {
            ((Model.IMailGoPG)this).Track.Debug("Startup ...");
        }

        protected override void OnShutdown()
        {
            ((Model.IMailGoPG)this).Track.Debug("Shutdown ...");
        }
    }
}
