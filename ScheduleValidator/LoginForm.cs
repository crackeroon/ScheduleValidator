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
            var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var settings = configFile.AppSettings.Settings;
            var name = this.name.Text;
            var pass = this.pass.Text;
            if (String.IsNullOrEmpty(name) || String.IsNullOrEmpty(pass))
            {
                MessageBox.Show("Пустые логин или пароль", "Ошибка!");
                return;
            }

            if (register)
            {
                if (settings[name] != null)
                {
                    MessageBox.Show("Этот пользователь уже зарегистрирован", "Ошибка!");
                    return;
                }
                settings.Add(name, pass);
                configFile.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
            }
            else
            {
                if (settings[name] == null)
                {
                    MessageBox.Show("Этот пользователь не зарегистрирован", "Ошибка!");
                    return;
                }
                if (pass != settings[name].Value)
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
