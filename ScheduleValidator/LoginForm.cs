using System;
using System.Configuration;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScheduleValidator
{
    public partial class LoginForm : Form
    {
        Form _Parent;
        public LoginForm(Form parent)
        {
            InitializeComponent();
            _Parent = parent;
        }

        void do_login(bool register = false)
        {
            if (this.name.Text == String.Empty || this.pass.Text == String.Empty)
            {
                MessageBox.Show("Пустые логин или пароль", "Ошибка!");
                return;
            }
            NameValueCollection sAll;
            sAll = ConfigurationManager.AppSettings;

            if (register)
            {
            } else
            {
                string pass = ConfigurationManager.AppSettings.Get(this.name.Text);
                if (String.IsNullOrEmpty(pass))
                {
                    MessageBox.Show("Этот пользователь не зарегистрирован", "Ошибка!");
                    return;
                }
                if (pass != this.pass.Text)
                {
                    Console.WriteLine(pass);
                    MessageBox.Show("Неправильный логин или пароль", "Ошибка!");
                    return;
                }
            }
            this.Hide();
            _Parent.ShowDialog();
            this.Close();
        }

        private void log_in_Click(object sender, EventArgs e)
        {
            this.do_login(false);
        }

        private void registration_Click(object sender, EventArgs e)
        {
            this.do_login(true);
        }
    }
}
