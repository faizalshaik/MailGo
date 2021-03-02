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
    public class MEmail : IEmail
    {
        #region Fields

        protected List<string> m_to = new List<string>();
        protected List<string> m_cc = new List<string>();
        protected string m_body = "";
        protected string m_sender = "";
        protected string m_subject = "";

        #endregion

        #region IEmail Members

        List<string> IEmail.TO
        {
            get { return this.m_to; }
        }

        List<string> IEmail.CC
        {
            get { return this.m_cc; }
        }

        string IEmail.Body
        {
            get { return this.m_body; }
            set { this.m_body = value; }
        }

        string IEmail.Sender
        {
            get { return this.m_sender; }
            set { this.m_sender = value; }
        }

        string IEmail.Subject
        {
            get { return this.m_subject; }
            set { this.m_subject = value; }
        }

        #endregion
    }
}
