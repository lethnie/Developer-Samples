using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ResponsesExportExcel
{
    public partial class FormLogin : Form
    {
        ClientInfo clientInfo = new ClientInfo();

        public FormLogin()
        {
            InitializeComponent();
            tbPassword.PasswordChar = '*';
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        }

        private void OK_Click(object sender, EventArgs e)
        {
            clientInfo.name = tbLogin.Text;
            clientInfo.password = tbPassword.Text;
            this.DialogResult = DialogResult.OK;
        }

        public ClientInfo getLoginInfo()
        {
            return clientInfo;
        }
    }
}
