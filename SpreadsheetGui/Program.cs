using System;
using System.Windows.Forms;

namespace SocialSpreadSheet
{
    class SpreadsheetGuiApplicationContext : ApplicationContext
    {

        // Number of open forms
        private int _formCount;

        private static SpreadsheetGuiApplicationContext _appContext;

        private SpreadsheetGuiApplicationContext()
        {
        }

        public static SpreadsheetGuiApplicationContext GetAppContext()
        {
            return _appContext ?? (_appContext = new SpreadsheetGuiApplicationContext());
        }

        /// <summary>
        /// Runs the form.
        /// </summary>
        /// <param name="form">Form to run.</param>
        public void RunForm(Form form)
        {
            // One more form running
            _formCount++;

            // Adds a method to the FormClosed event handler which simultaneously decrements
            // formCount and specified that if its new value is less than or equal to zero,
            // the thread should be exited.
            form.FormClosed += (o, e) => { if (--_formCount <= 0) ExitThread(); };

            // Actually runs the form.
            form.Show();
        }

    }


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

            SpreadsheetGuiApplicationContext appContext = SpreadsheetGuiApplicationContext.GetAppContext();

            appContext.RunForm(new Form1());
            Application.Run(appContext);

        }
    }
}
