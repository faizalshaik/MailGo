/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.22 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Management;
using System.Net.NetworkInformation;

namespace DataDesign.MailGO.License
{
    internal class Hardware
    {
        public static bool IsSame(Model.IHardware v_source, Model.IHardware v_target)
        {
            if (v_source == null || v_target == null) return false;
            /* 10:40 AM 2008/05/19 Kawane-san wanted not to check the MACAddress any more so I comment these code.
            if (v_source.MAC.Count != v_target.MAC.Count) return false;
            

            foreach (string t_src in v_source.MAC)
            {
                if (v_target.MAC.IndexOf(t_src) < 0) return false;
            }

            foreach (string t_tag in v_target.MAC)
            {
                if (v_source.MAC.IndexOf(t_tag) < 0) return false;
            }

             */
            
            if (v_source.CPU.Count != v_target.CPU.Count) return false;

            foreach (string t_src in v_source.CPU)
            {
                if (v_target.CPU.IndexOf(t_src) < 0) return false;
            }

            foreach (string t_tag in v_target.CPU)
            {
                if (v_source.CPU.IndexOf(t_tag) < 0) return false;
            }

            return true;
        }

        public static Model.IHardware Import(string v_source)
        {
            Model.IHardware t_target = new Model.MHardware();
            string[] t_fields = v_source.Split(new char[] { ':' });
            string t_maclist = t_fields[0];
            string t_cpulist = t_fields[1];

            t_fields = t_maclist.Split(new char[] { '.' });
            foreach (string t_mac in t_fields)
            {
                if (t_mac != "")
                {
                    t_target.MAC.Add(t_mac);
                }
            }

            t_fields = t_cpulist.Split(new char[] { '.' });
            foreach (string t_cpu in t_fields)
            {
                t_target.CPU.Add(t_cpu);
            }

            return t_target;
        }

        public static string Export(Model.IHardware v_hardware)
        {
            string t_maclist = "";
            string t_cpulist = "";

            foreach (string t_mac in v_hardware.MAC)
            {
                t_maclist += "." + t_mac;
            }
            if (t_maclist.Length > 0)
            {
                t_maclist = t_maclist.Substring(1);
            }

            foreach (string t_cpu in v_hardware.CPU)
            {
                t_cpulist += "." + t_cpu;
            }
            if (t_cpulist.Length > 0)
            {
                t_cpulist = t_cpulist.Substring(1);
            }

            return t_maclist + ":" + t_cpulist;
        }

        public static List<string> GetMACAddress()
        {
            List<string> t_target = new List<string>();
            /* 10:40 AM 2008/05/19 Kawane-san wanted not to check the MACAddress any more so I comment these code.
            NetworkInterface[] t_interface_list = NetworkInterface.GetAllNetworkInterfaces();

            foreach (NetworkInterface t_interface in t_interface_list)
            {
                switch (t_interface.NetworkInterfaceType)
                {
                    case NetworkInterfaceType.Ethernet:
                        t_target.Add(t_interface.GetPhysicalAddress().ToString());
                        break;
                    default:
                        break;
                }
            }
             */

            return t_target;
        }

        public static List<string> GetCPUId()
        {
            ManagementClass t_mc = new ManagementClass("Win32_Processor");
            ManagementObjectCollection t_moc = t_mc.GetInstances();
            List<string> t_target = new List<string>();

            foreach (ManagementObject t_obj in t_moc)
            {
                t_target.Add(t_obj.Properties["ProcessorId"].Value.ToString());
            }

            return t_target;
        }
    }
}
