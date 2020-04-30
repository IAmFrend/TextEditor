using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using word = Microsoft.Office.Interop.Word;
using Microsoft.Win32;

namespace Lab4
{
    public partial class MainForm : Form
    {
        //Переменная для сохранения файла
        string file_path ="";
        //Форма "о программе"
        Form InfoForm = new Form();
        //Цвета темы
        public Color mc, mpc, cc, tc;
        public MainForm()
        {
            InitializeComponent();
            //Установка формы (при запуске). Иначе - устанавливает серую форму.
            try
            {
                InterfaceOption interfaceOption = new InterfaceOption();
                RegistryKey txtRedOption = Registry.CurrentUser;
                RegistryKey Interface = txtRedOption.CreateSubKey("Interface");
                switch (Interface.GetValue("Theme").ToString())
                {
                    case ("0"):
                        mc = Color.FromArgb(255,193,7);
                        mpc = Color.FromArgb(255,243,80);
                        cc = Color.FromArgb(199, 145, 0);
                        tc = Color.Black;
                        break;
                    case ("1"):
                        mc = Color.FromArgb(55,71,79);
                        mpc = Color.FromArgb(98,114,123);
                        cc = Color.FromArgb(16,32,39);
                        tc = Color.White;
                        break;
                    case ("2"):
                        mc = Color.FromArgb(0,187,0);
                        mpc = Color.FromArgb(14,132,54);
                        cc = Color.FromArgb(0,192,100);
                        tc = Color.White;
                        break;
                    case ("3"):
                        mc = Color.White;
                        mpc = Color.White;
                        cc = Color.White;
                        tc = Color.Black;
                        break;
                    case ("4"):
                        mc = Color.FromArgb(213, 0, 0);
                        mpc = Color.FromArgb(255, 81, 49);
                        cc = Color.FromArgb(155, 0, 0);
                        tc = Color.White;
                        break;
                }
            }
            catch
            {
                InterfaceOption interfaceOption = new InterfaceOption();
                RegistryKey txtRedOption = Registry.CurrentUser;
                RegistryKey Interface = txtRedOption.CreateSubKey("Interface");
                Interface.SetValue("Theme", "3");
            }
            try
            {
                BackColor = mc;
            }
            catch
            {
                BackColor = Color.Gray;
                foreach (ToolStripComboBox micb 
                    in toolStrip1.Items.OfType<ToolStripComboBox>())
                {
                    micb.BackColor = Color.Gray;
                    micb.ForeColor = Color.Black;
                }
                foreach (ToolStripMenuItem micb
                    in menuStrip1.Items.OfType<ToolStripMenuItem>())
                {
                    micb.BackColor = Color.Gray;
                    micb.ForeColor = Color.Black;
                }
            }
            finally
            {
                ForeColor = tc;
                menuStrip1.BackColor = cc;
                menuStrip1.ForeColor = tc;
                statusStrip1.BackColor = cc;
                statusStrip1.ForeColor = tc;
                toolStrip1.BackColor = mpc;
                toolStrip1.ForeColor = tc;
                foreach (ToolStripMenuItem MItem 
                    in menuStrip1.Items.OfType<ToolStripMenuItem>())
                {
                    MItem.BackColor = cc;
                    MItem.ForeColor = tc;
                    foreach (ToolStripItem TItem 
                        in MItem.DropDownItems.OfType<ToolStripItem>())
                    {
                        TItem.BackColor = cc;
                        TItem.ForeColor = tc;
                    }
                    foreach (ToolStripSeparator TSep 
                        in MItem.DropDownItems.OfType<ToolStripSeparator>())
                    {
                        TSep.BackColor = cc;
                        TSep.ForeColor = tc;
                    }
                }
            }

        }

        //Заполнение шрифтов и их размеров при запуске
        private void MainForm_Load(object sender, EventArgs e)
        {
            toolStripComboBox2.SelectedIndex = 0;
            foreach (FontFamily font in FontFamily.Families)
            {
                toolStripComboBox1.Items.Add(font.Name);
            }
            toolStripComboBox1.SelectedIndex 
                = toolStripComboBox1.FindStringExact("Arial");
            Program.Font_Name = toolStripComboBox1.Text;
            Program.Font_Size = Convert.ToSingle(toolStripComboBox2.Text);
        }

        //Изменение даты
        private void timer1_Tick(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text 
                = DateTime.Now.ToLongDateString() 
                + " " + DateTime.Now.ToLongTimeString();
        }

        //Создание окна текста
        private void новыйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Redactor redactor = new Redactor();
            redactor.Form_Create("New_file", this);
        }

        //Изменеие шрифта
        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Program.Font_Name = toolStripComboBox1.Text;
            if (Program.RedactorTextBox != null)
            {
                Program.RedactorTextBox.SelectionFont 
                    = new Font(Program.Font_Name, 
                    Program.Font_Size, Program.style);
            }
        }

        //Изменение размера шрифта
        private void toolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            Program.Font_Size = Convert.ToSingle(toolStripComboBox2.Text);
            if (Program.RedactorTextBox != null)
            {
                Program.RedactorTextBox.SelectionFont = 
                    new Font(Program.Font_Name, 
                    Program.Font_Size, Program.style);
            }
        }

        //Смена выравнивания (по левому краю)
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            Program.RedactorTextBox.SelectionAlignment = HorizontalAlignment.Left;
        }

        //Смена выравнивания (по центру)
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            Program.RedactorTextBox.SelectionAlignment = HorizontalAlignment.Center;
        }

        //Смена выравнивания (по правому краю)
        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            Program.RedactorTextBox.SelectionAlignment = HorizontalAlignment.Right;
        }

        //Смена состояния стиля (жирный)
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            switch (toolStripButton1.Pressed)
            {
                case (true):
                    Program.style = Program.style ^ FontStyle.Bold;
                    break;
                case (false):
                    Program.style = Program.style | FontStyle.Bold;
                    break;
            }
            Program.RedactorTextBox.SelectionFont = 
                new Font(Program.Font_Name, Program.Font_Size, Program.style);
        }

        //Смена состояния стиля (курсив)
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            switch (toolStripButton2.Pressed)
            {
                case (true):
                    Program.style = Program.style ^ FontStyle.Italic;
                    break;
                case (false):
                    Program.style = Program.style | FontStyle.Italic;
                    break;
            }
            Program.RedactorTextBox.SelectionFont 
                = new Font(Program.Font_Name, Program.Font_Size, Program.style);
        }
        //Смена состояния стиля (Подчёркивание)
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            switch (toolStripButton3.Pressed)
            {
                case (true):
                    Program.style = Program.style ^ FontStyle.Underline;
                    break;
                case (false):
                    Program.style = Program.style | FontStyle.Underline;
                    break;
            }
            Program.RedactorTextBox.SelectionFont 
                = new Font(Program.Font_Name, Program.Font_Size, Program.style);
        }
        //закрытие активной тестовой формы
        private void закрытьФайлToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form form = ActiveMdiChild;
            form.Close();
        }

        //Закрытие программы
        private void выходИзПрограммыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //Процедура проверки закрытия программы
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            switch (MessageBox.Show("Зaкрыть программу?",
                "Текстовый редактор", 
                MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                case (DialogResult.Yes):
                    e.Cancel = false;
                    Application.ExitThread();
                    break;
                case (DialogResult.No):
                    e.Cancel = true;
                    break;
            }
        }

        //Открытие файла
        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        //Добавление нового текстового окна при открытии формы
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            if (openFileDialog1.FileName != "")
            {
                file_path = openFileDialog1.FileName;
                Redactor redactor = new Redactor();
                redactor.Form_Create(openFileDialog1.FileName, this);
                switch (openFileDialog1.FilterIndex)
                {
                    case (1):
                        if (File.Exists(openFileDialog1.FileName))
                        {
                            StreamReader reader = 
                                new StreamReader(openFileDialog1.FileName);
                            Program.RedactorTextBox.Text = reader.ReadToEnd();
                            reader.Close();
                        }
                        break;
                    case (2):
                        word.Application application = new word.Application();
                        word.Document documents 
                            = application.Documents.Open(openFileDialog1.FileName);
                        try
                        {
                            for (int i = 0; i < documents.Paragraphs.Count; ++i)
                            {
                                Program.RedactorTextBox.Font 
                                    = new Font(documents.Paragraphs[i + 1].Range.Font.Name, 
                                    documents.Paragraphs[i + 1].Range.Font.Size);
                                Program.RedactorTextBox.
                                    AppendText(documents.Paragraphs[i + 1].Range.Text.ToString());
                            }
                        }
                        catch
                        {

                        }
                        finally
                        {
                            documents.Close();
                            application.Quit();
                        }
                        break;
                    case (3):
                        word.Application application1 = new word.Application();
                        word.Document documents1 = application1.Documents.Open(openFileDialog1.FileName);
                        try
                        {
                            for (int i = 0; i < documents1.Paragraphs.Count; ++i)
                            {
                                Program.RedactorTextBox.Font = new Font(documents1.Paragraphs[i + 1].Range.Font.Name, documents1.Paragraphs[i + 1].Range.Font.Size);
                                Program.RedactorTextBox.AppendText(documents1.Paragraphs[i + 1].Range.Text.ToString());
                            }
                        }
                        catch
                        {

                        }
                        finally
                        {
                            documents1.Close();
                            application1.Quit();
                        }
                        break;
                }                
            }
            else
            {
                MessageBox.Show("Выберите файл", 
                    "Текстовый реадктор", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        //Сохранение по другому пути
        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Redactor redactor = new Redactor();
            redactor.saveFile.Filter =
                "Файл блокнота|*.txt|Microsoft Word 97-2003|*.doc|" +
                "Microsoft Word|*.docx";
            redactor.save_dialog_execute();
        }

        //Сохранение файла
        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(file_path);
            switch (file_path != "")
            {
                case (true):
                    if (File.Exists(file_path))
                    {
                        FileInfo fileInfo = new FileInfo(file_path);
                        switch (fileInfo.Extension)
                        {
                            case ("txt"):
                                StreamWriter writer = new StreamWriter(file_path);
                                writer.Write(Program.RedactorTextBox.Text);
                                writer.Close();
                                break;
                            case ("doc"):
                                word.Application MSOW97 = new word.Application();
                                word.Document document97 
                                    = MSOW97.Documents.Add(Visible: true);
                                word.Paragraph paragraph97 
                                    = document97.Paragraphs.Add();
                                paragraph97.Range.Text = Program.RedactorTextBox.Text;
                                paragraph97.Range.Font.Name = Program.Font_Name;
                                paragraph97.Range.Font.Size = Program.Font_Size;
                                document97.SaveAs2(file_path, 
                                    word.WdSaveFormat.wdFormatDocument97);
                                document97.Close();
                                MSOW97.Quit();
                                break;
                            case ("docx"):
                                word.Application MSOW = new word.Application();
                                word.Document document = MSOW.Documents.Add(Visible: true);
                                word.Paragraph paragraph = document.Paragraphs.Add();
                                paragraph.Range.Text = Program.RedactorTextBox.Text;
                                paragraph.Range.Font.Name = Program.Font_Name;
                                paragraph.Range.Font.Size = Program.Font_Size;
                                document.SaveAs2(file_path, 
                                    word.WdSaveFormat.wdFormatDocumentDefault);
                                document.Close();
                                MSOW.Quit();
                                break;
                        }
                    }
                    else
                    {
                        Redactor redactor = new Redactor();
                        redactor.save_dialog_execute();
                    }
                    break;
                case (false):

                    break;
            }
        }

        //Вывод статистики
        private void статистикаСимволовToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChartStaticstic chart = new ChartStaticstic();
            chart.Show();
        }

        //Смена состояния стиля (зачёркнутый)
        private void ToolStripButton7_Click(object sender, EventArgs e)
        {
            switch (toolStripButton7.Pressed)
            {
                case (true):
                    Program.style = Program.style ^ FontStyle.Strikeout;
                    break;
                case (false):
                    Program.style = Program.style | FontStyle.Strikeout;
                    break;
            }
            Program.RedactorTextBox.SelectionFont =
                new Font(Program.Font_Name, Program.Font_Size, Program.style);
        }

        //Шрифт +1
        private void ToolStripButton11_Click(object sender, EventArgs e)
        {
            try
            {
                if (toolStripComboBox2.SelectedIndex +1 <= toolStripComboBox2.Items.Count)
                toolStripComboBox2.SelectedIndex += 1;
            }
            catch
            { }
            toolStripComboBox2_SelectedIndexChanged(sender, e);
        }

        //Шрифт -1
        private void ToolStripButton12_Click(object sender, EventArgs e)
        {
            try
            {
                if (toolStripComboBox2.SelectedIndex - 1 > -1)
                    toolStripComboBox2.SelectedIndex -= 1;
            }
            catch
            { }
            toolStripComboBox2_SelectedIndexChanged(sender, e);
        }

        //Все прописные
        private void ToolStripButton13_Click(object sender, EventArgs e)
        {
            string output = "";
            foreach (char ch in Program.RedactorTextBox.SelectedText)
            {
                if (!(Char.IsUpper(ch)))
                {
                    output += Char.ToUpper(ch);
                }
                else
                {
                    output += ch;
                }
            }
            Program.RedactorTextBox.SelectedText = output;
        }

        //Все прописные
        private void ToolStripButton14_Click(object sender, EventArgs e)
        {
            string output = "";
            foreach (char ch in Program.RedactorTextBox.SelectedText)
            {
                if (!(Char.IsUpper(ch)))
                {
                    output += ch;
                }
                else
                {
                    output += Char.ToLower(ch);
                }
            }
            Program.RedactorTextBox.SelectedText = output;
        }

        //окно настройки интерфейса
        private void настройкиИнтерфейсаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InterfaceOption interfaceOption = new InterfaceOption();
            interfaceOption.Show(this);
        }

        //Вывод данных о программе
        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InfoForm.Height = 200;
            InfoForm.Width = 250;
            InfoForm.FormBorderStyle = FormBorderStyle.None;
            InfoForm.StartPosition = FormStartPosition.CenterScreen;
            InfoForm.BackColor = Color.Black;
            Panel panel = new Panel();
            panel.Dock = DockStyle.Bottom;
            panel.Height = 25;
            InfoForm.Controls.Add(panel);
            Button button = new Button();
            Label label = new Label();
            label.Dock = DockStyle.Fill;
            label.ForeColor = Color.White;
            label.Text = "Текстовый реадктор.\n " +
                "Используется в качестве учебных средств. " +
                "\n\n\nВ возможность входит:\n " +
                "работа с файлами форматов " +
                "(txt, Microsoft Word, Microsoft Exel, PDF). \n " +
                "Построение диаграмм символов. \n\n\n\n " +
                "Разработчик: Львов Михаил Дмитриевич";
            button.FlatStyle = FlatStyle.Flat;
            button.ForeColor = Color.White;
            button.Text = "Закрыть";
            button.Click += button_click;
            panel.Controls.Add(button);
            InfoForm.Controls.Add(label);
            InfoForm.ShowDialog();
            panel.Dispose();
        }

        //Закрытие формы данных о программе
        private void button_click(object sender, EventArgs e)
        {
            InfoForm.Close();
        }
    }
}
