/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.18 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;
using System.Windows.Forms;


namespace DataDesign.MailGO.Address
{
    internal class CSV
    {
        private string m_filename;
        private Encoding m_encoding = Encoding.UTF8;

        public Encoding Encoding
        {
            get { return m_encoding; }
            set { m_encoding = value; }
        }

        public CSV(string v_filename)
        {
            this.m_filename = v_filename;
        }

        public CSV(string v_filename, Encoding v_encoding)
        {
            this.m_filename = v_filename;
            this.m_encoding = v_encoding;
        }

        //public List<Model.IAddress> ReadShipmentData()
        //{
        //    Assembly _assembly = Assembly.GetExecutingAssembly();
        //    StreamReader t_reader = new StreamReader(_assembly.GetManifestResourceStream("DataDesign.MailGO.Address.Shipment.csv"));
        //    List<Model.IAddress> t_address_list = new List<Model.IAddress>();
        //    string t_line;
        //    Model.IAddress t_address;
        //
        //    while ((t_line = t_reader.ReadLine()) != null)
        //    {
        //        t_address = this.Read(t_line);
        //        if (t_address != null)
        //        {
        //            t_address_list.Add(t_address);
        //        }
        //    }
        //
        //    t_reader.Close();
        //    return t_address_list;
        //}        

        public List<Model.IAddress> Read()
        {
            //StreamReader t_reader = new StreamReader(File.OpenRead(this.m_filename), Encoding.GetEncoding("iso-2022-jp"));
            StreamReader t_reader = DataDesign.MailGO.Utils.EncodingTools.OpenTextFile(this.m_filename);            
            List<Model.IAddress> t_address_list = new List<Model.IAddress>();
            string t_line;
            Model.IAddress t_address;

            while ((t_line = t_reader.ReadLine()) != null)
            {
                t_address = this.Read(t_line);
                if (t_address != null)
                {
                    t_address_list.Add(t_address);
                }
            }

            if (t_address_list.Count > 0)
                this.m_encoding = t_reader.CurrentEncoding;            

            t_reader.Close();
            return t_address_list;
        }

        public void Write(List<Model.IAddress> v_address_list)
        {
            StreamWriter t_writer = new StreamWriter(File.Open(this.m_filename, FileMode.Create, FileAccess.Write), Encoding.UTF8);
            //StreamWriter t_writer = new StreamWriter(File.Open(this.m_filename, FileMode.Create, FileAccess.Write));            

            foreach (Model.IAddress t_address in v_address_list)
            {
                t_writer.WriteLine(string.Format("\"{0}\",\"{1}\"", this.DoubleQuote(t_address.Company), this.DoubleQuote(t_address.Suffix)));
            }

            t_writer.Close();
        }

        private string DoubleQuote(string v_source)
        {
            string t_target = "";
            char t_char = ' ';

            for (int t_idx = 0; t_idx < v_source.Length; t_idx++)
            {
                t_char = v_source[t_idx];
                if (t_char == '"')
                {
                    t_target += '"';
                }
                t_target += t_char;
            }

            return t_target;
        }

        private Model.IAddress Read(string v_line)
        {
            Model.IAddress t_address = null;
            List<string> t_field_list = new List<string>();
            this.ParseFields(ref t_field_list, v_line);

            if (t_field_list.Count == 2)
            {
                t_address = new Model.MAddress(t_field_list[0], t_field_list[1]);
            }
            t_field_list.Clear();

            return t_address;
        }

        private void ParseFields(ref List<string> v_field_list, string v_line)
        {
            int t_pos = -1;
            while (t_pos < v_line.Length)
            {
                v_field_list.Add(GetField(v_line, ref t_pos));
            }
        }

        private string GetField(string v_line, ref int v_start)
        {
            string t_field = "";
            int t_from = v_start + 1;
            int t_next;

            this.SkipBlank(v_line, ref t_from);

            if (t_from == v_line.Length - 1)
            {
                v_start = t_from + 1;
            }
            else
            {
                if (v_line[t_from] == '"')
                {
                    if (t_from == v_line.Length - 1)
                    {
                        t_from++;
                        t_field = "\"";
                    }
                    else
                    {
                        t_next = this.FindSingleQuote(v_line, t_from + 1);
                        t_field = v_line.Substring(t_from + 1, t_next - t_from - 1).Replace("\"\"", "\"");
                        this.SkipToComma(v_line, ref t_next);
                        v_start = t_next;
                    }
                }
                else
                {
                    t_next = v_line.IndexOf(',', t_from);
                    if (t_next < 0)
                    {
                        v_start = v_line.Length;
                        t_field = v_line.Substring(t_from);
                    }
                    else
                    {
                        v_start = t_next;
                        t_field = v_line.Substring(t_from, t_next - t_from);
                    }
                }
            }
            return t_field;
        }

        private void SkipToComma(string v_line, ref int v_start)
        {
            while (v_start < v_line.Length)
            {
                if (v_line[v_start] == ',') break;
                v_start++;
            }
        }

        private void SkipBlank(string v_line, ref int v_start)
        {
            while (v_start < v_line.Length)
            {
                if (v_line[v_start] != ' ') break;
                v_start++;
            }
        }

        private int FindSingleQuote(string v_line, int v_start)
        {
            int t_idx = v_start - 1;
            while (++t_idx < v_line.Length)
            {
                if (v_line[t_idx] == '"')
                {
                    if (t_idx < v_line.Length - 1 && v_line[t_idx + 1] == '"')
                    {
                        t_idx++;
                        continue;
                    }
                    else
                    {
                        return t_idx;
                    }
                }
            }
            return t_idx;
        }
    }
}
