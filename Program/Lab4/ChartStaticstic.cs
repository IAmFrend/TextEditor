using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Lab4
{
    public partial class ChartStaticstic : Form
    {
        //Данные статистики
        public static int Numbers, LatinLtr, KirilLtr, Space, Other;
        public ChartStaticstic()
        {
            InitializeComponent();
        }

        //Создание Excel-документа
        private void button2_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;
            Redactor redactor = new Redactor();
            redactor.Excel_Create(Program.Redact.Name, LatinLtr, KirilLtr, Numbers, Space, Other);
            button2.Enabled = true;
        }

        //закрытие формы
        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        //Появление статистики
        public void ChartStaticstic_Load(object sender, EventArgs e)
        {
            try
            {
                Text = "Символьная статистика " + Program.Redact.Text;
                LatinLtr = 0;
                KirilLtr = 0;
                Numbers = 0;
                Space = 0;
                Other = 0;
                foreach (char ch in Program.RedactorTextBox.Text)
                {
                    if ((ch >= 'a') && (ch <= 'z') || (ch >= 'A') && (ch <= 'Z'))
                    {
                        LatinLtr++;
                    }
                    else
                    if ((ch >= 'а') && (ch <= 'я') || (ch >= 'А') && (ch <= 'Я'))
                    {
                        KirilLtr++;
                    }
                    else
                    if ((ch >= '0') && (ch <= '9'))
                    {
                        Numbers++;
                    }
                    else
                    if (ch == ' ')
                    {
                        Space++;
                    }
                    else
                    {
                        Other++;
                    }
                }
                if (LatinLtr>0)
                chart1.Series["SymbStatic"].Points.AddXY("Латинские буквы " + LatinLtr.ToString(), LatinLtr);
                if (KirilLtr > 0)
                    chart1.Series["SymbStatic"].Points.AddXY("Русские буквы " + KirilLtr.ToString(), KirilLtr);
                if (Numbers > 0)
                    chart1.Series["SymbStatic"].Points.AddXY("Числа " + Numbers.ToString(), Numbers);
                if (Space > 0)
                    chart1.Series["SymbStatic"].Points.AddXY("Пробелы " + Space.ToString(), Space);
                if (Other > 0)
                    chart1.Series["SymbStatic"].Points.AddXY("Другие символы " + Other.ToString(), Other);
            }
            catch
            {
                MessageBox.Show("Невозможно вывести статистику");
                Close();
            }
        }
    }
}
