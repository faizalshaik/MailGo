using System;
using System.Xml;
using System.Reflection;
using System.IO;
using log4net;

namespace DataDesign.MailGO.Utils
{
    public class ConfigUtil
    {        
        public const string CONFIG_FILENAME = "options.xml";
        private static log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private ConfigUtil()
        { }

        public static string ReadSetting(string key)
        {            
            // load config document for current assembly
            XmlDocument doc = loadConfigDocument();

            // retrieve appSettings node
            XmlNode node = doc.SelectSingleNode("//appSettings");

            if (node == null)
            {
                logger.Error("appSettings section not found in config file.");
                return null;
            }

            try
            {
                // select the 'add' element that contains the key
                XmlElement elem = (XmlElement)node.SelectSingleNode(string.Format("//add[@key='{0}']", key));

                if (elem != null)
                {
                    return elem.Attributes["value"].Value;
                }
                else
                {
                    logger.Error("key not found in config file.");
                    return null;
                }
            }
            catch
            {
                throw;
            }
        }

        public static void WriteSetting(string key, string value)
        {
            // load config document for current assembly
            XmlDocument doc = loadConfigDocument();

            // retrieve appSettings node
            XmlNode node = doc.SelectSingleNode("//appSettings");

            if (node == null)
                logger.Error("appSettings section not found in config file.");

            try
            {
                // select the 'add' element that contains the key
                XmlElement elem = (XmlElement)node.SelectSingleNode(string.Format("//add[@key='{0}']", key));

                if (elem != null)
                {
                    // add value for key
                    elem.SetAttribute("value", value);
                }
                else
                {
                    // key was not found so create the 'add' element 
                    // and set it's key/value attributes 
                    elem = doc.CreateElement("add");
                    elem.SetAttribute("key", key);
                    elem.SetAttribute("value", value);
                    node.AppendChild(elem);
                }
                doc.Save(getConfigFilePath());
            }
            catch
            {
                throw;
            }
        }

        public static void RemoveSetting(string key)
        {
            // load config document for current assembly
            XmlDocument doc = loadConfigDocument();

            // retrieve appSettings node
            XmlNode node = doc.SelectSingleNode("//appSettings");

            try
            {
                if (node == null)
                    logger.Error("appSettings section not found in config file.");
                else
                {
                    // remove 'add' element with coresponding key
                    node.RemoveChild(node.SelectSingleNode(string.Format("//add[@key='{0}']", key)));
                    doc.Save(getConfigFilePath());
                }
            }
            catch (NullReferenceException e)
            {
                logger.Error(string.Format("The key {0} does not exist.", key), e);
            }
        }

        private static XmlDocument loadConfigDocument()
        {
            XmlDocument doc = null;
            try
            {
                doc = new XmlDocument();
                doc.Load(getConfigFilePath());
                return doc;
            }
            catch (System.IO.FileNotFoundException e)
            {
                logger.Error("No configuration file found.", e);
                return null;
            }
        }

        public static string getConfigFilePath()
        {
            return Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), CONFIG_FILENAME);
        }

        public static void ReconfigLogAppender(string configFilePath)
        {
            FileInfo fInfo = new FileInfo(configFilePath);
            log4net.Config.XmlConfigurator.Configure(fInfo);

            // reconfigure the RollingFileAppender for this assembly
            // We will look for the appender whose name matches with the item name (e.g. MyProject.MyAdmin.dll) of this assembly
            string assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string appenderNames = "RollingFileAppender";

            log4net.Appender.IAppender[] appenders = logger.Logger.Repository.GetAppenders();
            foreach (log4net.Appender.IAppender appender in appenders)
            {
                if (appenderNames.Equals(appender.Name))
                {
                    log4net.Appender.RollingFileAppender fileAppender = appender as log4net.Appender.RollingFileAppender;
                    fileAppender.File = Path.Combine(Path.GetDirectoryName(assemblyLocation), Path.GetFileName(fileAppender.File));
                    fileAppender.ActivateOptions();
                    break;
                }
            }
        }
    }
}
