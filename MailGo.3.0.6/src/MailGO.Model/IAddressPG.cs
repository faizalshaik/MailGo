/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.18 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;

namespace DataDesign.MailGO.Model
{
    public interface IAddressPG
    {
        IMailGoPG MailGo { get; }

        string UserDefinedFile { get; set; }
        //string ShipmentFile { get; set; }

        List<IAddress> UserDefinedList { get; }
        //List<IAddress> ShipmentList { get; }

        Encoding Encoding { get; set; }

        void CreateList();
        void CheckEmail(IEmail v_email, out bool v_cancel);
    }
}
