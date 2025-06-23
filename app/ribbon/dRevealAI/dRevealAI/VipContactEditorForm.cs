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
            lstContacts.DataSource = _contacts.ToList(); // Bind fresh copy
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

        //#region Designer Code (Auto-generated)

        //private Label label1;
        //private TextBox txtEmail;
        //private Button btnAdd;
        //private Button btnRemove;
        //private Button btnSave;
        //private ListBox lstContacts;
        //private Button btnCancel;

        //private void InitializeComponent()
        //{
        //    this.label1 = new System.Windows.Forms.Label();
        //    this.txtEmail = new System.Windows.Forms.TextBox();
        //    this.btnAdd = new System.Windows.Forms.Button();
        //    this.lstContacts = new System.Windows.Forms.ListBox();
        //    this.btnRemove = new System.Windows.Forms.Button();
        //    this.btnSave = new System.Windows.Forms.Button();
        //    this.btnCancel = new System.Windows.Forms.Button();
        //    this.SuspendLayout();

        //    // label1
        //    this.label1.AutoSize = true;
        //    this.label1.Location = new System.Drawing.Point(12, 9);
        //    this.label1.Name = "label1";
        //    this.label1.Size = new System.Drawing.Size(73, 13);
        //    this.label1.TabIndex = 0;
        //    this.label1.Text = "New Email:";

        //    // txtEmail
        //    this.txtEmail.Location = new System.Drawing.Point(91, 6);
        //    this.txtEmail.Name = "txtEmail";
        //    this.txtEmail.Size = new System.Drawing.Size(250, 20);
        //    this.txtEmail.TabIndex = 1;

        //    // btnAdd
        //    this.btnAdd.Location = new System.Drawing.Point(347, 4);
        //    this.btnAdd.Name = "btnAdd";
        //    this.btnAdd.Size = new System.Drawing.Size(75, 23);
        //    this.btnAdd.TabIndex = 2;
        //    this.btnAdd.Text = "Add";
        //    this.btnAdd.UseVisualStyleBackColor = true;
        //    this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);

        //    // lstContacts
        //    this.lstContacts.FormattingEnabled = true;
        //    this.lstContacts.Location = new System.Drawing.Point(15, 40);
        //    this.lstContacts.Name = "lstContacts";
        //    this.lstContacts.Size = new System.Drawing.Size(407, 290);
        //    this.lstContacts.TabIndex = 3;

        //    // btnRemove
        //    this.btnRemove.Location = new System.Drawing.Point(428, 40);
        //    this.btnRemove.Name = "btnRemove";
        //    this.btnRemove.Size = new System.Drawing.Size(75, 23);
        //    this.btnRemove.TabIndex = 4;
        //    this.btnRemove.Text = "Remove";
        //    this.btnRemove.UseVisualStyleBackColor = true;
        //    this.btnRemove.Click += new System.EventHandler(this.btnRemove_Click);

        //    // btnSave
        //    this.btnSave.Location = new System.Drawing.Point(347, 336);
        //    this.btnSave.Name = "btnSave";
        //    this.btnSave.Size = new System.Drawing.Size(75, 23);
        //    this.btnSave.TabIndex = 5;
        //    this.btnSave.Text = "Save";
        //    this.btnSave.UseVisualStyleBackColor = true;
        //    this.btnSave.Click += new System.EventHandler(this.btnSave_Click);

        //    // btnCancel
        //    this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        //    this.btnCancel.Location = new System.Drawing.Point(428, 336);
        //    this.btnCancel.Name = "btnCancel";
        //    this.btnCancel.Size = new System.Drawing.Size(75, 23);
        //    this.btnCancel.TabIndex = 6;
        //    this.btnCancel.Text = "Cancel";
        //    this.btnCancel.UseVisualStyleBackColor = true;

        //    // Form
        //    this.ClientSize = new System.Drawing.Size(515, 371);
        //    this.Controls.Add(this.btnCancel);
        //    this.Controls.Add(this.btnSave);
        //    this.Controls.Add(this.btnRemove);
        //    this.Controls.Add(this.lstContacts);
        //    this.Controls.Add(this.btnAdd);
        //    this.Controls.Add(this.txtEmail);
        //    this.Controls.Add(this.label1);
        //    this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
        //    this.MaximizeBox = false;
        //    this.MinimizeBox = false;
        //    this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        //    this.Text = "Manage VIP Contacts";
        //    this.ResumeLayout(false);
        //    this.PerformLayout();
        //}

        //#endregion

    }
}
