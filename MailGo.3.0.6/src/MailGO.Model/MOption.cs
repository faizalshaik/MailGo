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
    public class MOption : IOption
    {
        #region Fields

        protected bool m_check_all_address = true;
        protected bool m_check_only_first_address = false;
        protected bool m_check_cc_line = false;
        protected bool m_check_same_suffix = false;
        // 2010/9/18
        protected bool m_check_word = false;
        protected bool m_check_excel = false;
        protected bool m_check_powerpoint = false;
        protected bool m_check_text = false;
        //

        #endregion

        #region IOption Members

        bool IOption.CheckAllAddress
        {
            get { return this.m_check_all_address; }
            set 
            { 
                this.m_check_all_address = value;
                this.m_check_only_first_address = !value;
            }
        }

        bool IOption.CheckOnlyFirstAddress
        {
            get { return this.m_check_only_first_address; }
            set 
            { 
                this.m_check_only_first_address = value;
                this.m_check_all_address = !value;
            }
        }

        bool IOption.CheckCCLine
        {
            get { return this.m_check_cc_line; }
            set { this.m_check_cc_line = value; }
        }

        bool IOption.CheckSameSuffix
        {
            get { return this.m_check_same_suffix; }
            set { this.m_check_same_suffix = value; }
        }

        // 2010/9/18
        bool IOption.CheckWord
        {
            get { return this.m_check_word; }
            set { this.m_check_word = value; }
        }
        bool IOption.CheckExcel
        {
            get { return this.m_check_excel; }
            set { this.m_check_excel = value; }
        }
        bool IOption.CheckPowerPoint
        {
            get { return this.m_check_powerpoint; }
            set { this.m_check_powerpoint = value; }
        }
        bool IOption.CheckText
        {
            get { return this.m_check_text; }
            set { this.m_check_text = value; }
        }
        //

        #endregion
    }
}
