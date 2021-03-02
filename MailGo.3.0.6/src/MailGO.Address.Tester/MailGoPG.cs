/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.18 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.IO;
using System.Reflection;

namespace DataDesign.MailGO.Address.Tester
{
    internal class MailGoPG : Model.MMailGoPG
    {
        private const string c_lic_address = "awoekghaoskdghoawieghdkaoskdghaksdkghaslkdoekqyghdkgha";

        protected override void OnCreateTrack()
        {
            this.m_track = new Track();
        }

        protected override void OnCreateAddress()
        {
            Address.Share.License = c_lic_address;
            this.m_address = Address.Share.CreatePG(this);
        }

        protected override void OnStartup()
        {
            ((Model.IMailGoPG)this).Track.Debug("Startup ...");
            
            string t_shipment = ConfigurationManager.AppSettings["ShipmentList"].ToString();
            string t_user_defined = ConfigurationManager.AppSettings["UserDefinedList"].ToString();
            string t_path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            t_shipment = Path.Combine(t_path, t_shipment);
            t_user_defined = Path.Combine(t_path, t_user_defined);

            //((Model.IMailGoPG)this).Address.ShipmentFile = t_shipment;
            ((Model.IMailGoPG)this).Address.UserDefinedFile = t_user_defined;

            ((Model.IMailGoPG)this).Track.Debug(string.Format("Shipment Address List: {0}", t_shipment));
            ((Model.IMailGoPG)this).Track.Debug(string.Format("User Defined Address List: {0}", t_user_defined));
        }

        protected override void OnShutdown()
        {
            ((Model.IMailGoPG)this).Track.Debug("Shutdown ...");
        }
    }
}
