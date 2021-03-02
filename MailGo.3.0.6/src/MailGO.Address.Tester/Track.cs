/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.18 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;

namespace DataDesign.MailGO.Address.Tester
{
    internal class Track : Model.MTrack
    {
        private log4net.ILog m_log;

        public Track()
            : base()
        {
            this.m_log = log4net.LogManager.GetLogger(typeof(Track));
        }

        protected override void OnDebug(string v_msg)
        {
            this.m_log.Debug(v_msg);
        }

        protected override void OnError(Exception e)
        {
            this.m_log.Error(e);
        }

        protected override void OnDebug(Exception e)
        {
            this.m_log.Debug(e);
        }

        protected override void OnError(string v_msg)
        {
            this.m_log.Error(v_msg);
        }
    }
}
