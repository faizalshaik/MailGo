/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.21 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.Text;

namespace DataDesign.MailGO.License
{
    public class Share
    {
        public static Model.ILicensePG CreatePG(Model.IMailGoPG v_mailgo)
        {
            return new LicensePG(v_mailgo);
        }
    }
}
