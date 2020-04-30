using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace Lab4
{
    static class Program
    {
        //Количество форм текстовых файлов
        public static int New_Form_Count = 0;
        //Данные шрифта
        public static string Font_Name;
        public static float Font_Size;
        //Данные активной формы и поля ввода
        public static Form Redact;
        public static RichTextBox RedactorTextBox;
        public static FontStyle style;
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new AutorizForm());
        }
    }
}
