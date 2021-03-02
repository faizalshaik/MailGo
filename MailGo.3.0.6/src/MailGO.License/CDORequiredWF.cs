using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DataDesign.MailGO.License
{
    public partial class CDORequiredWF : Form
    {
        public CDORequiredWF(bool v_ol2k7)
        {
            InitializeComponent();

            this.lblMsg2K7.Visible = v_ol2k7;
            this.lblMsg2K3.Visible = !v_ol2k7;
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public static void Show(bool v_ol2k7)
        {
            CDORequiredWF t_form = new CDORequiredWF(v_ol2k7);
            t_form.ShowDialog();
        }
    }
}