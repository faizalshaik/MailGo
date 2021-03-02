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
    public interface IActivationPG
    {
        IMailGoPG MailGo { get; }

        IOption Option { get; }
        bool Status { get; set; }

        bool Activated { get; set; }

        bool ActivatedReally { get; set; }

        DateTime DeadLine { get; set; }

        string Language { get; set; }

        void SaveOption();
        void LoadOption();
        void ChangeOption();
        void BeforeSendEmail(IEmail v_email, out bool v_cancel);

        void CleanUp();
        void Activate();
        void DeActivate();
    }
}
