using System;
using System.Windows.Forms;

namespace MsgViewer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                Application.Run(new ViewerForm());
            }
            catch (Exception exception)
            {
                // ReSharper disable once LocalizableElement
                MessageBox.Show(exception.Message, "An exception occurred", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
