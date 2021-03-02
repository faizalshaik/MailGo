/* Copyright 2008 Data Design Vietnam. All rights reserved.
 * 
 * Created 2008.01.18 Tran Dinh Thoai
 * 
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DataDesign.MailGO.Address
{
    internal partial class AddressWF : Form
    {
        private Model.IAddressPG m_address;
        private bool m_new = false;

        public AddressWF(Model.IAddressPG v_address)
        {
            this.m_address = v_address;
            InitializeComponent();
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AddressWF_Load(object sender, EventArgs e)
        {
            //this.FillListView(this.lsvShipment, this.m_address.ShipmentList);
            this.FillListView(this.lsvUserDefined, this.m_address.UserDefinedList);
            this.SetViewMode();
        }

        private void FillListView(ListView v_listview, List<Model.IAddress> v_address_list)
        {
            Model.IAddress t_address;
            ListViewItem t_item;

            v_listview.Items.Clear();
            for (int t_idx = 0; t_idx < v_address_list.Count; t_idx++)
            {
                t_address = v_address_list[t_idx];
                t_item = new ListViewItem(t_address.Company);
                t_item.SubItems.Add(t_address.Suffix);
                v_listview.Items.Add(t_item);
            }
        }

        private void EnableEditor()
        {
            this.txtCompanyName.Enabled = true;
            this.txtSuffix.Enabled = true;
        }

        private void DisableEditor()
        {
            this.txtCompanyName.Enabled = false;
            this.txtSuffix.Enabled = false;
        }

        private void SetViewMode()
        {
            this.DisableEditor();

            this.cmdModify.Visible = true;
            this.cmdSave.Visible = false;
            this.cmdDelete.Visible = true;
            this.cmdDiscard.Visible = false;

            this.cmdNew.Enabled = true;
            this.cmdImport.Enabled = true;
            this.CheckViewMode();
        }

        private void CheckViewMode()
        {
            if (this.lsvUserDefined.Items.Count == 0 || this.lsvUserDefined.SelectedIndices.Count == 0)
            {
                this.cmdModify.Enabled = false;
                this.cmdDelete.Enabled = false;
            }
            else
            {
                this.cmdModify.Enabled = true;
                this.cmdDelete.Enabled = true;
            }
        }

        private void SetEditMode()
        {
            this.EnableEditor();

            this.cmdModify.Visible = false;
            this.cmdSave.Visible = true;
            this.cmdDelete.Visible = false;
            this.cmdDiscard.Visible = true;

            this.cmdNew.Enabled = false;
            this.cmdImport.Enabled = false;
        }

        private void ClearData()
        {
            this.txtCompanyName.Text = "";
            this.txtSuffix.Text = "";
        }

        private void FillData()
        {
            ListViewItem t_item = this.lsvUserDefined.SelectedItems[0];
            this.txtCompanyName.Text = t_item.SubItems[0].Text;
            this.txtSuffix.Text = t_item.SubItems[1].Text;
        }

        private void SaveData()
        {
            ListViewItem t_item = this.lsvUserDefined.SelectedItems[0];
            t_item.SubItems[0].Text = this.txtCompanyName.Text;
            t_item.SubItems[1].Text = this.txtSuffix.Text;
        }

        private void AddData()
        {
            ListViewItem t_item = new ListViewItem(this.txtCompanyName.Text);
            t_item.SubItems.Add(this.txtSuffix.Text);
            this.lsvUserDefined.Items.Add(t_item);
        }

        private void lsvUserDefined_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.lsvUserDefined.SelectedIndices.Count == 1)
            {
                this.FillData();
            }
            else
            {
                this.ClearData();
            }
            this.CheckViewMode();
        }

        private void cmdNew_Click(object sender, EventArgs e)
        {
            this.m_new = true;
            this.ClearData();
            this.SetEditMode();
        }

        private void cmdModify_Click(object sender, EventArgs e)
        {
            this.m_new = false;
            this.SetEditMode();
        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            if (this.m_new)
            {
                this.AddData();
                this.SetViewMode();
                this.lsvUserDefined.Items[this.lsvUserDefined.Items.Count - 1].Selected = true;
                this.lsvUserDefined_SelectedIndexChanged(null, null);
            }
            else
            {
                this.SaveData();
                this.SetViewMode();
                this.lsvUserDefined_SelectedIndexChanged(null, null);
            }
        }

        private void cmdDiscard_Click(object sender, EventArgs e)
        {
            this.SetViewMode();
            this.lsvUserDefined_SelectedIndexChanged(null, null);
        }

        private void cmdDelete_Click(object sender, EventArgs e)
        {
            if (this.lsvUserDefined.SelectedIndices.Count > 0)
            {
                this.lsvUserDefined.Items.RemoveAt(this.lsvUserDefined.SelectedIndices[0]);
            }
            if (this.lsvUserDefined.Items.Count > 0)
            {
                this.lsvUserDefined.Items[this.lsvUserDefined.Items.Count - 1].Selected = true;
            }
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            ListViewItem t_item;
            Model.IAddress t_address;

            this.m_address.UserDefinedList.Clear();
            for (int t_idx = 0; t_idx < this.lsvUserDefined.Items.Count; t_idx++)
            {
                t_item = this.lsvUserDefined.Items[t_idx];
                t_address = new Model.MAddress(t_item.SubItems[0].Text, t_item.SubItems[1].Text);
                this.m_address.UserDefinedList.Add(t_address);
            }

            CSV t_csv = new CSV(this.m_address.UserDefinedFile, this.m_address.Encoding);
            t_csv.Write(this.m_address.UserDefinedList);
            this.Close();
        }

        private void cmdImport_Click(object sender, EventArgs e)
        {
            this.ctlOpenFileDialog.FileName = "";
            this.ctlOpenFileDialog.ShowDialog(this);

            if (this.ctlOpenFileDialog.FileName == "") return;

            CSV t_csv = new CSV(this.ctlOpenFileDialog.FileName);
            List<Model.IAddress> t_list = t_csv.Read();

            foreach (Model.IAddress t_address in t_list)
            {
                ListViewItem t_item = new ListViewItem(t_address.Company);
                t_item.SubItems.Add(t_address.Suffix);
                this.lsvUserDefined.Items.Add(t_item);
            }
        }

        private void cmdExport_Click(object sender, EventArgs e)
        {
            this.ctlSaveFileDialog.FileName = "";
            this.ctlSaveFileDialog.Filter = "CSV|*.csv";
            this.ctlSaveFileDialog.AddExtension = true;
            this.ctlSaveFileDialog.ShowDialog(this);

            if (this.ctlSaveFileDialog.FileName == "") return;

            CSV w_csv = new CSV(this.ctlSaveFileDialog.FileName);
            CSV r_csv = new CSV(this.m_address.UserDefinedFile, this.m_address.Encoding);
            w_csv.Write(r_csv.Read());

        }

    }
}