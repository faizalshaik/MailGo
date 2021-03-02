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
    public class MAddressPG : IAddressPG
    {
        #region Fields

        protected IMailGoPG m_mailgo;
        protected string m_user_defined_file = "";
        protected string m_shipment_file = "";
        protected List<IAddress> m_user_defined_list = new List<IAddress>();
        //protected List<IAddress> m_shipment_list = new List<IAddress>();
        protected Encoding m_encoding = Encoding.Default;        

        #endregion

        #region Constructors

        public MAddressPG(IMailGoPG v_mailgo)
        {
            this.m_mailgo = v_mailgo;
        }

        #endregion

        #region IAddressPG Members

        IMailGoPG IAddressPG.MailGo
        {
            get { return this.m_mailgo; }
        }

        string IAddressPG.UserDefinedFile
        {
            get { return this.m_user_defined_file; }
            set 
            { 
                this.m_user_defined_file = value;
                this.OnLoadUserDefinedList();
            }
        }

        //string IAddressPG.ShipmentFile
        //{
        //    get { return this.m_shipment_file; }
        //    set 
        //    { 
        //        this.m_shipment_file = value;
        //        this.OnLoadShipmentList();
        //    }
        //}

        Encoding IAddressPG.Encoding
        {
            get { return m_encoding; }
            set { m_encoding = value; }
        }

        List<IAddress> IAddressPG.UserDefinedList
        {
            get { return this.m_user_defined_list; }
        }

        //List<IAddress> IAddressPG.ShipmentList
        //{
        //    get { return this.m_shipment_list; }
        //}

        void IAddressPG.CreateList()
        {
            this.OnCreateList();
        }

        void IAddressPG.CheckEmail(IEmail v_email, out bool v_cancel)
        {
            this.OnCheckEmail(v_email, out v_cancel);
        }

        #endregion

        #region Overridable

        protected virtual void OnCreateList() { }

        protected virtual void OnCheckEmail(IEmail v_email, out bool v_cancel) 
        {
            v_cancel = false;
        }

        //protected virtual void OnLoadShipmentList() { }

        protected virtual void OnLoadUserDefinedList() { }

        #endregion

    }
}
