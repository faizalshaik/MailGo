/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.18 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace DataDesign.MailGO.Address
{
    internal class AddressPG : Model.MAddressPG
    {
        public AddressPG(Model.IMailGoPG v_mailgo) : base(v_mailgo) { }

        protected override void OnCreateList()
        {
            //this.OnLoadShipmentList();
            this.OnLoadUserDefinedList();

            AddressWF t_form = new AddressWF(this);
            t_form.ShowDialog();

            //this.OnLoadShipmentList();
            this.OnLoadUserDefinedList();
        }

        //protected override void OnLoadShipmentList()
        //{
        //    CSV t_csv = new CSV(this.m_shipment_file);
        //    List<Model.IAddress> t_address_list = t_csv.ReadShipmentData();
        //
        //    this.m_shipment_list.Clear();
        //    for (int t_idx = 0; t_idx < t_address_list.Count; t_idx++)
        //    {
        //        this.m_shipment_list.Add(t_address_list[t_idx]);
        //    }
        //}

        protected override void OnLoadUserDefinedList()
        {
            if (this.m_user_defined_file == null || this.m_user_defined_file == "")
                return;

            if(File.Exists(m_user_defined_file)==false)
                return;

            CSV t_csv = new CSV(this.m_user_defined_file);
            List<Model.IAddress> t_address_list = t_csv.Read();
            this.m_encoding = t_csv.Encoding;

            this.m_user_defined_list.Clear();
            for (int t_idx = 0; t_idx < t_address_list.Count; t_idx++)
            {
                this.m_user_defined_list.Add(t_address_list[t_idx]);
            }
        }

        protected override void OnCheckEmail(DataDesign.MailGO.Model.IEmail v_email, out bool v_cancel)
        {
            Checker t_checker = new Checker(this);
            t_checker.Run(v_email, out v_cancel);
        }
    }
}
