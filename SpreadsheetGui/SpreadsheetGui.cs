using System;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using SS;
using SocialSpreadSheet;
using SpreadsheetUtilities;
using System.Net;

namespace SocialSpreadSheet
{
    public partial class Form1 : Form
    {
        private Spreadsheet _spreadsheet;
        private string _filename;

        /// <summary>
        ///     Opens a new spreadsheet GUI window.
        /// </summary>
        public Form1()
        {
            InitializeComponent();

            _spreadsheet = null;
            _filename = "Not Connected";

            Text = "Spreadsheet Program - " + _filename;

            // Registers my displaySelection method.
            spreadsheetPanel1.SelectionChanged += DisplaySelection;
            spreadsheetPanel1.SetSelection(0, 0);
        }

        /// <summary>
        ///     Opens a new spreadsheet GUI window.
        /// </summary>
        /// <param name="filename">File to be opened.</param>
        public Form1(String server, int port, string filename, string password, bool isCreate)
        {
            InitializeComponent();
            try
            {
                _spreadsheet = new SocialSpreadSheet.Spreadsheet(server, port, filename, password, isCreate, isValid, normalize, "ps6");
                _filename = filename;

                // Registers my displaySelection method.
                spreadsheetPanel1.SelectionChanged += DisplaySelection;
                spreadsheetPanel1.SetSelection(0, 0);
               
                RegisterHandlers();
            }
            catch (SpreadsheetReadWriteException ex)
            {
                MessageBox.Show("Spreadsheet could not open. " + ex.Message + " Please close this spreadsheet window.",
                                "Open Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public override sealed string Text
        {
            get { return base.Text; }
            set { base.Text = value; }
        }

        /// <summary>
        /// </summary>
        /// <param name="ss">The SpreadsheetPanel in the current Form1</param>
        private void DisplaySelection(SpreadsheetPanel ss)
        {
            if (_spreadsheet == null) return;
            int col, row;
            ss.GetSelection(out col, out row);
            // Convert col, row index to spreadsheet cell names.
            var cellName = ((char) (col + 65)) + (row + 1).ToString(CultureInfo.InvariantCulture);

            // Displays selected cell's name
            CellNameBox.Invoke(new Action(() => { CellNameBox.Text = cellName; }));
            var content = _spreadsheet.GetCellContents(cellName);
            // If content is a formula, prepend "=" before displaying
            var f = content as Formula;
            if (f != null)
            {
                ContentBox.Invoke(new Action(() => { ContentBox.Text = "=" + f; }));
            }
                // Otherwise just display the content.
            else
            {
                ContentBox.Invoke(new Action(() => { ContentBox.Text = content.ToString(); }));
            }
            // No need to fetch the value from the spreadsheet again, just copy it from
            // the spreadsheetpanel. This avoids reworking the FormulaError message.
            string value;
            ss.GetValue(col, row, out value);
            ValueBox.Invoke(new Action(() => { ValueBox.Text = value; }));
        }

        private void DisplayUpdate(SpreadsheetPanel ss)
        {
            if (_spreadsheet == null) return;
            int col, row;
            ss.GetSelection(out col, out row);
            // Convert col, row index to spreadsheet cell names.
            var cellName = ((char)(col + 65)) + (row + 1).ToString(CultureInfo.InvariantCulture);

            // Displays selected cell's name
            CellNameBox.Invoke(new Action(() => { CellNameBox.Text = cellName; }));
            // No need to fetch the value from the spreadsheet again, just copy it from
            // the spreadsheetpanel. This avoids reworking the FormulaError message.
            string value;
            ss.GetValue(col, row, out value);
            ValueBox.Invoke(new Action(() => { ValueBox.Text = value; }));
        }

        /// <summary>
        ///     Ensures that only valid cell names are used. (One letter followed by a number from 1 to 99).
        /// </summary>
        /// <param name="s">Cell name to be validated.</param>
        /// <returns>True if s is a valid cell name, otherwise false.</returns>
        private bool isValid(string s)
        {
            return Regex.IsMatch(s, "^[a-zA-Z][1-9][0-9]?$");
        }


        /// <summary>
        ///     Normalizes all cell names to upper case for case insensitive comparison.
        /// </summary>
        /// <param name="s">Cell name to be normalized.</param>
        /// <returns>A new, all uppercase version of s.</returns>
        private string normalize(string s)
        {
            return s.ToUpper();
        }

        /// <summary>
        ///     Handles saving once a file name has been entered and confirmed.
        /// </summary>
        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            var fileName = saveFileDialog1.FileName;

            // Warns user that a file already exists and will be overwritten. Asks user to confirm the action.
            if (File.Exists(fileName) && fileName != _filename)
            {
                var result =
                    MessageBox.Show(
                        "The file, \"" + Path.GetFileName(fileName) +
                        ",\" already exists.\nThe save operation will overwrite this file, and all previous information will be lost."
                        + "\nAre you sure you want to continue?", "Continue save?", MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    // Saves if user selected yes.
                    try
                    {
                        _spreadsheet.Save(fileName);
                        _filename = fileName;
                        Text = "Spreadsheet program - " + _filename;
                    }
                    catch (SpreadsheetReadWriteException ex)
                    {
                        MessageBox.Show("Spreadsheet did not save. " + ex.Message, "Save Error", MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                    }
                }
            }
                // If entered filename is this spreadsheet's original filename, or if filename does not already
                // exist, this will save without comment.
            else
            {
                try
                {
                    _spreadsheet.Save(fileName);
                    _filename = fileName;
                    Text = "Spreadsheet program - " + _filename;
                }
                catch (SpreadsheetReadWriteException ex)
                {
                    MessageBox.Show("Spreadsheet did not save. " + ex.Message, "Save Error", MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                }
            }
        }

        private void ContentBox_KeyDown(object sender, KeyEventArgs e)
        {
            // If key pressed isn't enter or return, quit. Otherwise sets cell contents.
            if (e.KeyCode != Keys.Return && e.KeyCode != Keys.Enter)
            {
                return;
            }
            ISet<string> toRedisplay = new HashSet<string>();

            try
            {
                _spreadsheet.Change(CellNameBox.Text, ContentBox.Text);
            }
            catch (FormulaFormatException ex)
            {
                // Creates a popup message box explaining error to user.
                MessageBox.Show(ex.Message, "Invalid Formula", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Sets ContentBox.Text back to cell's content.
                var content = _spreadsheet.GetCellContents(CellNameBox.Text);
                var f = content as Formula;
                if (f != null)
                {
                    ContentBox.Text = "=" + f;
                }
                else
                {
                    ContentBox.Text = content.ToString();
                }
            }
            catch (CircularException)
            {
                // Explains that a circular reference is not allowed.
                MessageBox.Show(
                    "Circular references are not permitted.\nA circular reference is created when a cell in a formula depends on itself for value, e.g." +
                    " when a cell in a formula refers to itself, or when a cell refers to another cell which itself refers to the original cell.",
                    "Circular Reference!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Redisplays contents of cell
                var content = _spreadsheet.GetCellContents(CellNameBox.Text);
                var f = content as Formula;
                if (f != null)
                {
                    ContentBox.Text = "=" + f;
                }
                else
                {
                    ContentBox.Text = content.ToString();
                }
            }

            // To silence the annoying beep! And probably for other reasons, too.
            e.SuppressKeyPress = true;
        }

        private void viewHelpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Displays my not so beautiful help file.
            Help.ShowHelp(this, @"..\..\..\Help.htm");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_spreadsheet == null) return;
            _spreadsheet.Leave();
        }

        private void createToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Connect("Create", this).Show();
        }

        private void joinToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Connect("Join", this).Show();
        }

        public void joinSpreadsheet(string server, int port, string filename, string password, bool isCreate)
        {
            
            if (_spreadsheet == null)
            {
                _spreadsheet = new Spreadsheet(server, port, filename, password, isCreate, isValid, normalize, "ps6");
                _filename = filename;
                RegisterHandlers();
                return;
            }
            // Something's already open.
            SpreadsheetGuiApplicationContext.GetAppContext().RunForm(new Form1(server, port, filename, password, isCreate));
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _spreadsheet.Save();
        }

        private void leaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _spreadsheet.Undo();
        }

        private void YouAreHosedEventHandler(object sender, EventArgs e)
        {
            MessageBox.Show("There was an error with your connection.\nPlease try reconnecting.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            CleanSpreadsheet();
        }

        private void FailedEventHandler(object sender, FailedEventArgs message)
        {
            MessageBox.Show("Your request failed. " + message.Message.Last(), "Failure!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            if (message.Message.First().StartsWith("CREATE") || message.Message.First().StartsWith("JOIN"))
            {
                DeregisterHandlers();
                _spreadsheet = null;
            }
            
        }

        private void JoinedEventHandler(object sender, EventArgs ignored)
        {
            var toDisplay = _spreadsheet.GetNamesOfAllNonemptyCells();


            foreach (var cellName in toDisplay)
            {
                // Converts the row portion of a cell name to a zero based int index for the
                // spreadsheetPanel.
                var row = Int32.Parse(cellName.Substring(1)) - 1;
                // Converts the column portion of a cell name to a zero based int index for the
                // spreadsheetPanel.
                var col = cellName[0] - 65;

                // Populates the spreadsheetPanel with the values of all non-empty cells.
                var val = _spreadsheet.GetCellValue(cellName);
                if (val is FormulaError)
                {
                    spreadsheetPanel1.SetValue(col, row, "FormulaError. " + ((FormulaError)val).Reason);
                }
                else
                {
                    spreadsheetPanel1.SetValue(col, row, val.ToString());
                }
            }

            // Displays cell name, value and content of default selection.
            DisplaySelection(spreadsheetPanel1);

            this.Invoke(new Action(() => { Text = "Spreadsheet Program - " + _filename; }));
        }

        private void UpdatedEventHandler(object sender, UpdatedEventArgs eventArgs)
        {
            foreach (var cellName in eventArgs.Cells)
            {
                // This updates the SpreadsheetPanel object with the values of all dependents of cellname.

                // Converts a cell name to col, row zero-based index.
                var row = Int32.Parse(cellName.Substring(1)) - 1;
                var col = cellName[0] - 65;

                var val = _spreadsheet.GetCellValue(cellName);
                if (val is FormulaError)
                {
                    spreadsheetPanel1.SetValue(col, row, "FormulaError. " + ((FormulaError)val).Reason);
                }
                else
                {
                    spreadsheetPanel1.SetValue(col, row, val.ToString());
                }
                DisplayUpdate(spreadsheetPanel1);
            }
        }

        private void SocketExceptionHandler(object sender, ErrorEventArgs error)
        {
            MessageBox.Show("The connection to the spreadsheet has been lost.\nPlease try to join again.\n\n" + error.GetException().Message, "Connection Lost", MessageBoxButtons.OK, MessageBoxIcon.Error);
            CleanSpreadsheet();
        }

        private void UndoEndEventHandler(object sender, EventArgs ignored)
        {
            MessageBox.Show("There are no more actions to Undo.", "Cannot Undo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GenericErrorEventHandler(object sender, EventArgs ignored)
        {
            MessageBox.Show("Well, this is embarrassing.\nIt looks like you have a big problem.\nPlease contact your software vendor.", "Unknown Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ConnectionClosedEventHandler(object sender, EventArgs ignored)
        {
            MessageBox.Show("The connection to the spreadsheet has closed.\nPlease rejoin.", "Connection Closed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            CleanSpreadsheet();
        }

        private void CleanSpreadsheet()
        {
            DeregisterHandlers();
            _spreadsheet = null;
            spreadsheetPanel1.Invoke(new Action(() => { spreadsheetPanel1.Clear(); }));
            CellNameBox.Invoke(new Action(() => { CellNameBox.Text = string.Empty; }));
            ValueBox.Invoke(new Action(() => { ValueBox.Text = string.Empty; }));
            ContentBox.Invoke(new Action(() => { ContentBox.Text = string.Empty; }));
            this.Invoke(new Action(() => { this.Text = "Spreadsheet Program - Not Connected"; }));
        }

        private void RegisterHandlers()
        {
            menuStrip1.Invoke(new Action(() => { saveToolStripMenuItem.Enabled = true; leaveToolStripMenuItem.Enabled = true; undoToolStripMenuItem.Enabled = true;
            ContentBox.Enabled = true; createToolStripMenuItem.Enabled = false; joinToolStripMenuItem.Enabled = false; }));

            _spreadsheet.Failed += FailedEventHandler;
            _spreadsheet.Joined += JoinedEventHandler;
            _spreadsheet.Updated += UpdatedEventHandler;
            _spreadsheet.SocketException += SocketExceptionHandler;
            _spreadsheet.UndoEnd += UndoEndEventHandler;
            _spreadsheet.YouAreHosed += YouAreHosedEventHandler;
            _spreadsheet.GenericError += GenericErrorEventHandler;
            _spreadsheet.ConnectionClosed += ConnectionClosedEventHandler;
        }

        private void DeregisterHandlers()
        {
            menuStrip1.Invoke(new Action(() => { saveToolStripMenuItem.Enabled = false; leaveToolStripMenuItem.Enabled = false;
            undoToolStripMenuItem.Enabled = false; ContentBox.Enabled = false; createToolStripMenuItem.Enabled = true; joinToolStripMenuItem.Enabled = true; }));
            _spreadsheet.Failed -= FailedEventHandler;
            _spreadsheet.Joined -= JoinedEventHandler;
            _spreadsheet.Updated -= UpdatedEventHandler;
            _spreadsheet.SocketException -= SocketExceptionHandler;
            _spreadsheet.UndoEnd -= UndoEndEventHandler;
            _spreadsheet.YouAreHosed -= YouAreHosedEventHandler;
            _spreadsheet.GenericError -= GenericErrorEventHandler;
            _spreadsheet.ConnectionClosed -= ConnectionClosedEventHandler;
        }

    }
}