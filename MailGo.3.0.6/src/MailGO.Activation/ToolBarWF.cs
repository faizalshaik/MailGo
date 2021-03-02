/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.23 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DataDesign.MailGO.Activation
{
    internal partial class ToolBarWF : Form
    {
        public static ToolBarWF sm_instance = null;

        public static ToolBarWF Instance
        {
            get
            {
                if (sm_instance == null)
                {
                    sm_instance = new ToolBarWF();
                }
                return sm_instance;
            }
        }

        public ToolBarWF()
        {
            InitializeComponent();
        }
    }
}