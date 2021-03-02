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
    public interface IEmail
    {
        string Sender { get; set; }
        List<string> TO { get; }
        List<string> CC { get; }
        string Body { get; set; }
        string Subject { get; set; }
    }
}
