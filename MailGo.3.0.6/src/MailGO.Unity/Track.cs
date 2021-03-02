/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.24 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;
using System.Windows.Forms;

namespace DataDesign.MailGO.Unity
{
    internal class Track : Model.MTrack
    {
        private log4net.ILog m_log;

        public Track()
            : base()
        {
            this.m_log = log4net.LogManager.GetLogger(typeof(Track));
            /*
            log4net.Appender.FileAppender rootAppender = (log4net.Appender.FileAppender)((log4net.Repository.Hierarchy.Hierarchy)log4net.LogManager.GetRepository()).Root.Appenders[0];
            string filename = rootAppender.File;
            MessageBox.Show(filename);
            */
        }

        protected override void OnDebug(string v_msg)
        {
            v_msg = DateTime.Now.ToString() + " : " + v_msg;            
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
            v_msg = DateTime.Now.ToString() + " : " + v_msg;
            this.m_log.Error(v_msg);
        }

        public static void Configure(out string v_config_file)
        {
            v_config_file = AppDomain.CurrentDomain.BaseDirectory + "MailGO.dll.config";
            log4net.Config.XmlConfigurator.Configure(new FileInfo(v_config_file));

            log4net.Appender.IAppender[] t_appender_list = log4net.LogManager.GetRepository().GetAppenders();
            foreach (log4net.Appender.IAppender t_appender in t_appender_list)
            {
                if (t_appender.GetType().ToString().Equals("log4net.Appender.RollingFileAppender"))
                {
                    log4net.Appender.RollingFileAppender t_file_appender = t_appender as log4net.Appender.RollingFileAppender;
                    t_file_appender.File = Path.Combine(Path.GetDirectoryName(v_config_file), Path.GetFileName(t_file_appender.File));
                    t_file_appender.ActivateOptions();
                }
            }
        }
    }
}
