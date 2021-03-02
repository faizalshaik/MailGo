/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.21 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.IO;
using System.Reflection;

namespace DataDesign.MailGO.License.Tester
{
    internal class MailGoPG : Model.MMailGoPG
    {
        private const string c_lic_license = "aoekghoakwhgodkahgoekhgaowighke1234dfhe@343tdogh&68*964546dhgoqdkgha";

        protected override void OnCreateTrack()
        {
            this.m_track = new Track();
        }

        protected override void OnCreateLicense()
        {
            License.Share.License = c_lic_license;
            this.m_license = License.Share.CreatePG(this);
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
