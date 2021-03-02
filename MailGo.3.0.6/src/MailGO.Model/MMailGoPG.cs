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
    public class MMailGoPG : IMailGoPG
    {
        #region Fields

        protected ITrack m_track;
        protected IAddressPG m_address;
        protected IActivationPG m_activation;
        #endregion

        #region IMailGoPG Members

        ITrack IMailGoPG.Track
        {
            get
            {
                if (this.m_track == null)
                {
                    this.OnCreateTrack();
                }
                return this.m_track;
            }
        }

        IAddressPG IMailGoPG.Address
        {
            get 
            {
                if (this.m_address == null)
                {
                    this.OnCreateAddress();
                }
                return this.m_address;
            }
        }

        bool IMailGoPG.Activated
        {
            get
            {
                return this.OnCheckActivated();
            }
        }

        IActivationPG IMailGoPG.Activation
        {
            get 
            {
                if (this.m_activation == null)
                {
                    this.OnCreateActivation();
                }
                return this.m_activation;
            }
        }

        void IMailGoPG.Startup()
        {
            this.OnStartup();
        }

        void IMailGoPG.Shutdown()
        {
            this.OnShutdown();
        }

        #endregion

        #region Overridable

        protected virtual bool OnCheckActivated()
        {
            return false;
        }

        protected virtual void OnCreateTrack()
        {
            this.m_track = new MTrack();
        }

        protected virtual void OnCreateAddress() 
        {
            this.m_address = new MAddressPG(this);
        }

        protected virtual void OnCreateActivation()
        {
            this.m_activation = new MActivationPG(this);
        }

        protected virtual void OnStartup() { }

        protected virtual void OnShutdown() { }

        #endregion
    }
}
