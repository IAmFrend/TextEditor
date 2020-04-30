using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using word = Microsoft.Office.Interop.Word;
using exel = Microsoft.Office.Interop.Excel;
using pdf = iTextSharp.text.pdf;
using pdffile = iTextSharp.text;

namespace Lab4
{

    class Redactor
    {
        //Создаваемая форма текстового поля
        public Form TextForm = new Form();
        //Элементы создаваемой формы
        public RichTextBox textBox = new RichTextBox();
        public SaveFileDialog saveFile = new SaveFileDialog();
        public StatusStrip statusStrip = new StatusStrip();
        public ToolStripLabel stripItemSymb = new ToolStripLabel();
        public ToolStripLabel stripItemWords = new ToolStripLabel();
        public ToolStripButton stripButton = new ToolStripButton();
        //Строка перезаписи для сохранения
        public string FileText = "";
        //Создание формы и размещение на ней компонентов.
        public void Form_Create(string Form_name, Form parent)
        {
            Program.New_Form_Count++;
            TextForm.Name = "TextForm_" + Program.New_Form_Count;
            TextForm.Text = Form_name + " " + Program.New_Form_Count;
            TextForm.Icon = parent.Icon;
            TextForm.MdiParent = parent;
            TextForm.TopMost = true;
            textBox.Name = "textBox_" + Program.New_Form_Count;
            textBox.Dock = DockStyle.Fill;
            textBox.Font = new Font(Program.Font_Name, Program.Font_Size, FontStyle.Regular);
            FileText = textBox.Text;
            stripItemSymb.Text = "Символы 0";
            stripItemWords.Text = "Слов 0";
            stripButton.Image = Properties.Resources.pdf;
            stripButton.Click += stripButton_Click;
            TextForm.Controls.Add(textBox);
            TextForm.Controls.Add(statusStrip);
            statusStrip.Items.Add(stripItemSymb);
            statusStrip.Items.Add(stripItemWords);
            statusStrip.Items.Add(stripButton);

            TextForm.FormClosing += TextForm_Closing;
            TextForm.Enter += TextFormChildEnter;

            textBox.KeyPress += textBox_KeyPress;
            TextForm.Show();
        }

        //Изменение полей статистики
        private void textBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            int word = 1;
            stripItemSymb.Text = "Символы " + textBox.Text.Length.ToString();
            switch (textBox.Text == "")
            {
                case (true):
                    word = 0;
                    break;
                case (false):
                    for (int i = 1; i < textBox.Text.Length; i++)
                    {
                        try
                        {
                            if ((textBox.Text[i] == ' ') && (textBox.Text[i + 1] != ' '))
                            {
                                word++;
                            }
                        }
                        catch
                        {

                        }
                    }
                    break;
            }
            stripItemWords.Text = "Слов " + word;
        }

        //Смена рабочей области при переходе на данную форму редактора
        private void TextFormChildEnter(object sender, EventArgs e)
        {
            Program.Redact = TextForm;
            Program.RedactorTextBox = textBox;
            Program.Font_Name = textBox.SelectionFont.FontFamily.Name;
            Program.Font_Size = textBox.SelectionFont.Size;
            FileText = textBox.Text;
        }

        //Закрытие формы текста. Диалог и сохранение при согласии пользователя
        private void TextForm_Closing(object sender, FormClosingEventArgs e)
        {
            switch (FileText != textBox.Text)
            {
                case (true):
                    saveFile.Filter = "Файл блокнота|*.txt|Microsoft Word 97-2003|*.doc|Microsoft Word|*.docx";
                    switch (MessageBox.Show("Файл изменён.\nСохранить файл " + TextForm.Text + "?", "Текстовый редактор", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question))
                    {
                        case (DialogResult.Yes):
                            e.Cancel = false;
                            save_dialog_execute();
                            Program.New_Form_Count--;
                            FileText = "";
                            break;
                        case (DialogResult.No):
                            e.Cancel = false;
                            Program.New_Form_Count--;
                            FileText = "";
                            break;
                        case (DialogResult.Cancel):
                            e.Cancel = true;
                            FileText = "";
                            break;
                    }
                    break;
                case (false):
                    e.Cancel = false;
                    FileText = "";
                    Program.New_Form_Count--;
                    break;
            }

        }

        //Флаг конвертирования файла в PDF
        private void stripButton_Click(object sender, EventArgs e)
        {
            saveFile.Filter = "Файл PDF|*.PDF";
            save_dialog_execute();
        }

        //Передача данных сохранения в окно сохранения
        public void save_dialog_execute()
        {
            saveFile.FileName = TextForm.Text;
            saveFile.FileOk += saveFile_OKClick;
            saveFile.ShowDialog();
        }

        //Сохранение файла в заданном формате 
        private void saveFile_OKClick(object sender, EventArgs e)
        {
            switch (saveFile.Filter == "Файл PDF|*.PDF")
            {
                case (true):
                    using (MemoryStream memory = new MemoryStream())
                    {
                        pdffile.Document document = new pdffile.Document(pdffile.PageSize.A4.Rotate());
                        pdf.BaseFont baseFont = pdf.BaseFont.CreateFont(@"C:\Users\IAmFriend\Desktop\ПО ОН\Лабы ООП\arial.ttf", pdf.BaseFont.IDENTITY_H, pdf.BaseFont.NOT_EMBEDDED);
                        pdffile.Font font = new pdffile.Font(baseFont, pdffile.Font.DEFAULTSIZE, pdffile.Font.NORMAL);
                        pdf.PdfWriter pdfWriter = pdf.PdfWriter.GetInstance(document, memory);
                        document.Open();
                        document.Add(new pdffile.Paragraph(Program.RedactorTextBox.Text, font));
                        document.Close();
                        byte[] content = memory.ToArray();
                        using (FileStream stream = File.Create(saveFile.FileName))
                        {
                            stream.Write(content, 0, (int)content.Length);
                        }
                    }
                    break;
                case (false):
                    if (saveFile.FileName != "")
                    {
                        switch (saveFile.FilterIndex)
                        {
                            case (1):
                                StreamWriter writer = new StreamWriter(saveFile.FileName);
                                writer.Write(Program.RedactorTextBox.Text);
                                writer.Close();
                                break;
                            case (2):
                                word.Application MSOW97 = new word.Application();
                                word.Document document97 = MSOW97.Documents.Add(Visible: true);
                                word.Paragraph paragraph97 = document97.Paragraphs.Add();
                                paragraph97.Range.Text = textBox.Text;
                                paragraph97.Range.Font.Name = Program.Font_Name;
                                paragraph97.Range.Font.Size = Program.Font_Size;
                                document97.SaveAs2(saveFile.FileName, word.WdSaveFormat.wdFormatDocument97);
                                document97.Close();
                                MSOW97.Quit();
                                break;
                            case (3):
                                word.Application MSOW = new word.Application();
                                word.Document document = MSOW.Documents.Add(Visible: true);
                                word.Paragraph paragraph = document.Paragraphs.Add();
                                paragraph.Range.Text = textBox.Text;
                                paragraph.Range.Font.Name = Program.Font_Name;
                                paragraph.Range.Font.Size = Program.Font_Size;
                                document.SaveAs2(saveFile.FileName, word.WdSaveFormat.wdFormatDocumentDefault);
                                document.Close();
                                MSOW.Quit();
                                break;
                        }
                    }
                    break;
            }
        }

        //Вывод статистики в лист Eхcel
        public void Excel_Create(string file_name, int latin, int kiril, int num, int spc, int oth)
        {
            exel.Application application = new exel.Application();
            exel.Workbook workbook = application.Workbooks.Add();
            exel.Worksheet worksheet = (exel.Worksheet)workbook.ActiveSheet;
            worksheet.Name = file_name;
            worksheet.Cells[1, 1] = "Латинские буквы";
            worksheet.Cells[2, 1] = "Кириллица";
            worksheet.Cells[3, 1] = "Числа";
            worksheet.Cells[4, 1] = "Пробелы";
            worksheet.Cells[5, 1] = "Прочее";
            worksheet.Cells[1, 2] = latin;
            worksheet.Cells[2, 2] = kiril;
            worksheet.Cells[3, 2] = num;
            worksheet.Cells[4, 2] = spc;
            worksheet.Cells[5, 2] = oth;
            worksheet.Columns.AutoFit();
            exel.ChartObjects chartObjects = (exel.ChartObjects)worksheet.ChartObjects(Type.Missing);
            exel.ChartObject chartObject = chartObjects.Add(50, 50, 250, 250);
            exel.Chart chart = chartObject.Chart;
            exel.SeriesCollection seriesCollection = (exel.SeriesCollection)chart.SeriesCollection(Type.Missing);
            exel.Series series = seriesCollection.NewSeries();
            chart.ChartType = exel.XlChartType.xl3DPie;
            series.XValues = worksheet.get_Range("A1", "A5");
            series.Values = worksheet.get_Range("B1", "B5");
            workbook.SaveAs();
            workbook.Close();
            application.Quit();
        }
    }
}
