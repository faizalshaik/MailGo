/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.18 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;

namespace DataDesign.MailGO.Address
{
    public class Share
    {

        public static Model.IAddressPG CreatePG(Model.IMailGoPG v_mailgo)
        {
            return new AddressPG(v_mailgo);
        }
    }
}
