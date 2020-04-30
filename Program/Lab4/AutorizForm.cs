using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime;
using Microsoft.Win32;

namespace Lab4
{
    public partial class AutorizForm : Form
    {
        public AutorizForm()
        {
            InitializeComponent();
        }

        //Кнопка закрытия
        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //Кнопка проверки
        private void Button2_Click(object sender, EventArgs e)
        {
            RegistryKey txtRedOption = Registry.CurrentUser;
            RegistryKey Autoriz = txtRedOption.CreateSubKey("Autorization");
            if (Autoriz.GetValue("Code").ToString() == textBox1.Text)
            {
                Autoriz.SetValue("Allow", "true");
                MainForm mainForm = new MainForm();
                mainForm.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Неправильно введён код авторизации", "Авторизация", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        //Сокрытие, если код уже введён
        private void AutorizForm_Shown(object sender, EventArgs e)
        {
            RegistryKey txtRedOption = Registry.CurrentUser;
            RegistryKey Autoriz = txtRedOption.CreateSubKey("Autorization");
            try
            {
                if (Autoriz.GetValue("Allow").ToString() == "true")
                {
                    MainForm mainForm = new MainForm();
                    mainForm.Show();
                    this.Hide();
                }
            }
            catch
            {

            }
            try
            {
                Autoriz.GetValue("Code").ToString();
            }
            catch
            {
                Autoriz.SetValue("Code", "12345");
            }
        }
    }
}
