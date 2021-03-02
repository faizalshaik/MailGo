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
    public class MTrack : ITrack
    {
        #region ITrack Members

        void ITrack.Error(Exception e)
        {
            this.OnError(e);
        }

        void ITrack.Error(string v_msg)
        {
            this.OnError(v_msg);
        }

        void ITrack.Debug(Exception e)
        {
            this.OnDebug(e);
        }

        void ITrack.Debug(string v_msg)
        {
            this.OnDebug(v_msg);
        }

        #endregion

        #region Overridable

        protected virtual void OnError(Exception e) { }

        protected virtual void OnError(string v_msg) { }

        protected virtual void OnDebug(Exception e) { }

        protected virtual void OnDebug(string v_msg) { }

        #endregion
    }
}
