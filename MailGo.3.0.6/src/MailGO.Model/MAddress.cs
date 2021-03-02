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
    public class MAddress : IAddress
    {
        #region Fields

        protected string m_company;
        protected string m_suffix;

        #endregion

        #region Construction

        public MAddress()
        {
            this.m_company = "";
            this.m_suffix = "";
        }

        public MAddress(string v_company, string v_suffix)
        {
            this.m_company = v_company;
            this.m_suffix = v_suffix;
        }

        #endregion

        #region IAddress Members

        string IAddress.Company
        {
            get { return this.m_company; }
            set { this.m_company = value; }
        }

        string IAddress.Suffix
        {
            get { return this.m_suffix; }
            set { this.m_suffix = value; }
        }

        IAddress IAddress.Clone()
        {
            return this.OnClone();
        }

        #endregion

        #region Overridable

        protected virtual IAddress OnClone()
        {
            return new MAddress(this.m_company, this.m_suffix);
        }

        #endregion
    }
}
