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
    public interface IOption
    {
        bool CheckAllAddress { get; set; }
        bool CheckOnlyFirstAddress { get; set; }
        bool CheckCCLine { get; set; }
        bool CheckSameSuffix { get; set; }
        // 2010/9/15
        bool CheckWord { get; set; }
        bool CheckExcel { get; set; }
        bool CheckPowerPoint { get; set; }
        bool CheckText { get; set; }
        //
    }
}
