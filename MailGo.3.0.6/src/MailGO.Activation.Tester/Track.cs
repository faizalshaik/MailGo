/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.23 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;

namespace DataDesign.MailGO.Activation.Tester
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

        public static void Configure()
        {
            string t_config_file = Assembly.GetExecutingAssembly().Location + ".config";
            log4net.Config.XmlConfigurator.Configure(new FileInfo(t_config_file));

            log4net.Appender.IAppender[] t_appender_list = log4net.LogManager.GetRepository().GetAppenders();
            foreach (log4net.Appender.IAppender t_appender in t_appender_list)
            {
                if (t_appender.GetType().ToString().Equals("log4net.Appender.RollingFileAppender"))
                {
                    log4net.Appender.RollingFileAppender t_file_appender = t_appender as log4net.Appender.RollingFileAppender;
                    t_file_appender.File = Path.Combine(Path.GetDirectoryName(t_config_file), Path.GetFileName(t_file_appender.File));
                    t_file_appender.ActivateOptions();
                }
            }
        }
    }
}
