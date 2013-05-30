using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SocialSpreadSheet
{
    public partial class Connect : Form
    {
        private Form1 _caller;
        public Connect(string action, Form1 caller)
        {
            InitializeComponent();
            this.Text = action;
            _caller = caller;
            submitButton.Text = action;
        }

        private void button1_Click(object sender, EventArgs e)
        {   
            int port;
            if (!Int32.TryParse(portBox.Text, out port))
            {
                port = 1984;
            }
            _caller.joinSpreadsheet(serverTextBox.Text, port, filenameTextBox.Text, passwordTextBox.Text, this.Text == "Create");
            this.Close();
        }

        private void passwordTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            // If key pressed isn't enter or return, quit. Otherwise sets cell contents.
            if (e.KeyCode != Keys.Return && e.KeyCode != Keys.Enter)
            {
                return;
            }
            button1_Click(this, new EventArgs());
            // To silence the annoying beep! And probably for other reasons, too.
            e.SuppressKeyPress = true;
        }

    }
}
