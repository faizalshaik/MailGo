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
    public class XMailGo : ApplicationException { }
    public class XAddress : XMailGo { }
    public class XLicense : XMailGo { }
    public class XCheckLicense : XLicense { }
    public class XInstallLicense : XLicense { }
    public class XActivateLicense : XLicense { }
    public class XBadActivationID : XActivateLicense { }
}
