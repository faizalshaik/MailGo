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
    public interface ITrack
    {
        void Error(Exception e);
        void Error(string v_msg);
        void Debug(Exception e);
        void Debug(string v_msg);
    }
}
