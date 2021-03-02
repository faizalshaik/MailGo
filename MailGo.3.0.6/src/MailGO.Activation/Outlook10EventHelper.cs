using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using OL = Microsoft.Office.Interop.Outlook;

namespace DataDesign.MailGO.Activation
{

    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch),
   GuidAttribute("0006300E-0000-0000-C000-000000000046")]
    public interface IOutlookApplicationEvents_10
    {
        [DispId(0x0000F002)]
        void ItemSend(object Item, ref bool Cancel);        
    }

    [ComVisible(true)]
    public class Outlook10EventHelper : IDisposable, IOutlookApplicationEvents_10
    {       
        private IConnectionPoint m_oConnectionPoint;
        private Int32 m_Cookie;
        private ActivationPG m_ActivePG;

        public Outlook10EventHelper()
		{
			m_oConnectionPoint = null;
			m_Cookie = 0;           
		}

        public void SetupConnection(OL.Application application, ActivationPG activatePG)
        {
            if (m_Cookie != 0) return;

            m_ActivePG = activatePG;

            // GUID of the DIID_ApplicationEvents dispinterface.
            Guid guid = new Guid("{0006300E-0000-0000-C000-000000000046}");
 
            // QI for IConnectionPointContainer.
            IConnectionPointContainer oConnectionPointContainer = (IConnectionPointContainer)application;

            try
            {
                // Find the connection point and then advise.
                m_ActivePG.MailGo.Track.Debug("Advise START");
                oConnectionPointContainer.FindConnectionPoint(ref guid, out m_oConnectionPoint);
                m_ActivePG.MailGo.Track.Debug("Advise CONTINUE");                
                m_oConnectionPoint.Advise(this, out m_Cookie);

                m_ActivePG.MailGo.Track.Debug("Advise FINISH");
            }            
            catch (Exception ex)
            {
                m_ActivePG.MailGo.Track.Debug("ERROR");
                m_ActivePG.MailGo.Track.Error(ex);
            }
        }

        public void RemoveConnection()
        {
            if (m_Cookie != 0)
            {
                m_oConnectionPoint.Unadvise(m_Cookie);
                m_oConnectionPoint = null;
                m_Cookie = 0;
            }
        }

        #region IDisposable Members

        public void Dispose()
        {
            RemoveConnection();
        }

        #endregion

        #region IOutlookApplicationEvents_10 Members

        [DispId(0x0000F002)] 
        public void ItemSend(object Item, ref bool Cancel)
        {
            m_ActivePG.Outlook_ItemSend(Item, ref Cancel);
        }
        #endregion
    }
}
