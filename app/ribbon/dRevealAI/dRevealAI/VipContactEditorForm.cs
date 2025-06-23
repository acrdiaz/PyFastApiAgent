using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace dRevealAI
{
    public partial class VipContactEditorForm: Form
    {

        private readonly List<string> _contacts = new List<string>();

        public VipContactEditorForm()
        {
            InitializeComponent();
            Load += OnLoadForm;
        }

        private void OnLoadForm(object sender, EventArgs e)
        {
            _contacts.AddRange(Properties.Settings.Default.VipContacts.Cast<string>());
            RefreshList();
        }

        private void RefreshList()
        {
            lstContacts.DataSource = null;
            // Bind fresh copy
            lstContacts.DataSource = _contacts.ToList();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            string input = txtEmail.Text.Trim();
            if (!string.IsNullOrWhiteSpace(input) && !_contacts.Contains(input))
            {
                _contacts.Add(input);
                Properties.Settings.Default.VipContacts.Add(input);
                Properties.Settings.Default.Save();
                RefreshList();
                txtEmail.Clear();
            }
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (lstContacts.SelectedItem != null)
            {
                string selected = lstContacts.SelectedItem.ToString();
                _contacts.Remove(selected);
                Properties.Settings.Default.VipContacts.Clear();
                Properties.Settings.Default.VipContacts.AddRange(_contacts.ToArray());
                Properties.Settings.Default.Save();
                RefreshList();
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.VipContacts.Clear();
            Properties.Settings.Default.VipContacts.AddRange(_contacts.ToArray());
            Properties.Settings.Default.Save();
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
