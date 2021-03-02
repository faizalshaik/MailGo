using System;
using System.Collections;
using System.Resources;
using System.Reflection;

namespace DataDesign.MailGO.Utils
{
    public class ResourceUtil
    {
        private static ResourceManager _resManager = null;        
        private ResourceUtil()
        {
        }
        public static string AssemblyName = string.Empty;
        public static ResourceManager Instance
        {
            get {
                if (_resManager == null)
                    return GetInstance();
                return _resManager;                    
            }
        }
        // This funciton to make Assembly.GetCallingAssembly() = this assembly!
        private static ResourceManager GetResourceManager(string resourceBaseName)
        {            
            // Get assembly thanks to a type object representing a class in this assembly
            AssemblyName = Assembly.GetAssembly(typeof(ResourceUtil)).FullName;
            return GetResourceManager(resourceBaseName, Assembly.GetAssembly(typeof(ResourceUtil)));
        }

        private static ResourceManager GetResourceManager(string resourceBaseName, Assembly assembly)
        {
            // check the cache             
            if (_resManager == null)          
                _resManager = new ResourceManager(resourceBaseName, assembly);
            return _resManager;
        }
        private static ResourceManager GetInstance()
        {
            string resourceBaseName = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.ToString();            
            return ResourceUtil.GetResourceManager(resourceBaseName);
        }

        public string GetString(string name)
        {
            return _resManager.GetString(name, System.Globalization.CultureInfo.CurrentCulture);
        }
    }
}
