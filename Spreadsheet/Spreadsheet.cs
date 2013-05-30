using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using System.Xml;
using CustomNetworking;
using SpreadsheetUtilities;
using System.Net;
using System.Text;
using System.IO;

namespace SocialSpreadSheet
{

    public class FailedEventArgs : EventArgs
    {
        public FailedEventArgs(List<string> message)
        {
            Message = message;
        }
        public List<string> Message { get; private set; }
    }

    public class UpdatedEventArgs : EventArgs
    {
        public UpdatedEventArgs(ISet<string> cells)
        {
            Cells = cells;
        }
        public ISet<string> Cells { get; private set; }
    }

    /// <summary>
    ///     An Spreadsheet object represents the state of a simple spreadsheet. 
    ///     A spreadsheet consists of an infinite number of named cells.
    ///     A string is a cell name if and only if it consists of one or more letters,
    ///     followed by one or more digits AND it satisfies the predicate IsValid.
    ///     For example, "A15", "a15", "XY032", and "BC7" are cell names so long as they
    ///     satisfy IsValid.  On the other hand, "Z", "X_", and "hello" are not cell names,
    ///     regardless of IsValid.
    ///     Any valid incoming cell name, whether passed as a parameter or embedded in a formula,
    ///     must be normalized with the Normalize method before it is used by or saved in
    ///     this spreadsheet.  For example, if Normalize is s => s.ToUpper(), then
    ///     the Formula "x3+a5" should be converted to "X3+A5" before use.
    ///     A spreadsheet contains a cell corresponding to every possible cell name.
    ///     In addition to a name, each cell has a contents and a value.  The distinction is
    ///     important.
    ///     The contents of a cell can be (1) a string, (2) a double, or (3) a Formula.  If the
    ///     contents is an empty string, we say that the cell is empty.  (By analogy, the contents
    ///     of a cell in Excel is what is displayed on the editing line when the cell is selected.)
    ///     In a new spreadsheet, the contents of every cell is the empty string.
    ///     The value of a cell can be (1) a string, (2) a double, or (3) a FormulaError.
    ///     (By analogy, the value of an Excel cell is what is displayed in that cell's position
    ///     in the grid.)
    ///     If a cell's contents is a string, its value is that string.
    ///     If a cell's contents is a double, its value is that double.
    ///     If a cell's contents is a Formula, its value is either a double or a FormulaError,
    ///     as reported by the Evaluate method of the Formula class.  The value of a Formula,
    ///     of course, can depend on the values of variables.  The value of a variable is the
    ///     value of the spreadsheet cell it names (if that cell's value is a double) or
    ///     is undefined (otherwise).
    ///     Spreadsheets are never allowed to contain a combination of Formulas that establish
    ///     a circular dependency.  A circular dependency exists when a cell depends on itself.
    ///     For example, suppose that A1 contains B1*2, B1 contains C1*2, and C1 contains A1*2.
    ///     A1 depends on B1, which depends on C1, which depends on A1.  That's a circular
    ///     dependency.
    /// </summary>
    public class Spreadsheet : AbstractSpreadsheet
    {
        // Representation invariants:
        // A cell is a string which must consist of one or more letters
        // followed by a non-zero digit followed by one or more digits 0-9.

        // Abstraction function:
        // Each key in cells is the name of a non-empty cell in a spreadsheet.
        // The corresponding value in cells is a cell whose Contents member 
        // corresponds to the contents of that spreadsheet cell. A valid cell
        // name that is not in cells refers to an empty cell whose contents
        // are an empty string.

        // The member dependencies is simply a helper to keep track of which
        // cells supply a value that another cell depends upon to calculate
        // its own value.

        private readonly Regex _validCellNameRegex = new Regex("^[a-zA-Z]+[0-9]+$");
        private readonly Dictionary<string, Cell> _cells;
        private readonly DependencyGraph _dependencies;

        private String _name;
        private String _version;

        private TcpClient _client;
        private StringSocket _socket;

        private string _currentChangeCell;
        private string _currentChangeContent;

        /// <summary>
        ///     True if this spreadsheet has been modified since it was created or saved
        ///     (whichever happened most recently); false otherwise.
        /// </summary>
        public override bool Changed { get; protected set; }

        public delegate void FailedEventHandler(object sender, FailedEventArgs eventArgs);
        public event FailedEventHandler Failed;

        public delegate void JoinedEventHandler(object sender, EventArgs ignored);
        public event JoinedEventHandler Joined;

        public delegate void UpdatedEventHandler(object sender, UpdatedEventArgs eventArgs);
        public event UpdatedEventHandler Updated;

        public delegate void SocketExceptionEventHandler(object sender, ErrorEventArgs eventArgs);
        public event SocketExceptionEventHandler SocketException;

        public delegate void UndoEndEventHandler(object sender, EventArgs ignored);
        public event UndoEndEventHandler UndoEnd;

        public delegate void YouAreHosedEventHandler(object sender, EventArgs ignored);
        public event YouAreHosedEventHandler YouAreHosed;

        public delegate void GenericErrorEventHandler(object sender, EventArgs ignored);
        public event GenericErrorEventHandler GenericError;

        public delegate void ConnectionClosedEventHandler(object sender, EventArgs ignored);
        public event ConnectionClosedEventHandler ConnectionClosed;

        /// <summary>
        /// </summary>
        /// <param name="filepath">Filepath of saved spreadsheet to load.</param>
        /// <param name="isValid">Used to impose additional requirements on the validity of cell names.</param>
        /// <param name="normalize">Used to force normalization of cell names.</param>
        /// <param name="version">
        ///     Provides a name for this version of a Spreadsheet (to correspond to the
        ///     particular combination of isValid and normalize parameters.
        /// </param>
        public Spreadsheet(string server, int port, string filename, string password, bool isCreate, Func<string, bool> isValid, Func<string, string> normalize, string version)
            : base(isValid, normalize, version)
        {

            _cells = new Dictionary<string, Cell>();
            _dependencies = new DependencyGraph();

            _client = new TcpClient();

            _client.BeginConnect(server, port, ConnectCallback, new Tuple<string, string, bool>(filename, password, isCreate));

        }

        /// <summary>
        /// Callback method called when connection to server is established.
        /// Sets up the string socket and raises the connected event.
        /// </summary>
        /// <param name="result">The IAsyncResult</param>
        private void ConnectCallback(IAsyncResult result)
        {
            try {
                _client.EndConnect(result);
            }
            catch (SocketException)
            {
            }
            // Creates a StringSocket from the TcpClient.
            _socket = new StringSocket(_client.Client, Encoding.UTF8);

            Tuple<string, string, bool> args = (Tuple<string, string, bool>)result.AsyncState;

            if (args.Item3)
            {
                Create(args.Item1, args.Item2);
                return;
            }
            // We're not creating; we're joining.
            Join(args.Item1, args.Item2);
        }

        /// <summary>
        /// Method called when a newly connected Spreadsheet wants to join a
        /// spreadsheet file on a server. The spreadsheet must be connected.
        /// </summary>
        /// <param name="name">The name of the spreadsheet file to join</param>
        /// <param name="password">The password of the spreadsheet file</param>
        public void Join(String name, String password)
        {
            _socket.BeginSend("JOIN\nName:" + name + "\nPassword:" + password + "\n", (exception, payload) => { }, null);
            _socket.BeginReceive(JoinCallback, null);
        }

        /// <summary>
        /// Call back to handle server response from JOIN command. If successful, leads to XmlCallback.
        /// </summary>
        /// <param name="response">The message from the server</param>
        /// <param name="exception">Any socket exceptions</param>
        /// <param name="state">A state object</param>
        private void JoinCallback(String response, Exception exception, object state)
        {
            if (exception != null || response == null)
            {
                SocketException(this, new ErrorEventArgs(exception));
                return;
            }
            if (response == null)
            {
                ConnectionClosed(this, null);
            }
            List<String> responses;
            if (response.StartsWith("JOIN "))
            {
                responses = new List<String>();
                responses.Add(response);
                _socket.BeginReceive(JoinCallback, responses);
                return;
            }
            if (response.Equals("ERROR"))
            {
                GenericError(this, null);
                return;
            }
            responses = (List<String>)state;
            responses.Add(response);
            if (responses.First() == "JOIN OK")
            {
                _name = response;
                _socket.BeginReceive(XmlCallback, responses);
                return;
            }
            _socket.BeginReceive(FailCallback, responses);
            
        }

        /// <summary>
        /// A callback for all fail message from server.
        /// </summary>
        /// <param name="response">The server response</param>
        /// <param name="exception">Any socket exception</param>
        /// <param name="state">State object</param>
        private void FailCallback(string response, Exception exception, object state)
        {
            if (exception != null)
            {
                SocketException(this, new ErrorEventArgs(exception));
                return;
            }
            if (response == null)
            {
                ConnectionClosed(this, null);
            }
            List<string> responses = (List<string>)state;
            responses.Add(response);
            if (response.StartsWith("Name:"))
            {
                // Calls itself to get the error message.
                _socket.BeginReceive(FailCallback, responses);
                return;
            }
            // If we get here, response is the error message.
            _currentChangeCell = null;
            _currentChangeContent = null;
            Failed(this, new FailedEventArgs(responses));
            _socket.BeginReceive(CallbackOperatorCallback, null);
        }

        /// <summary>
        /// Handles the end of the server response to the JOIN command. Parses the xml data for the spreadsheet.
        /// </summary>
        /// <param name="response">The server message</param>
        /// <param name="exception">Any socket exception</param>
        /// <param name="state">A state object</param>
        private void XmlCallback(String response, Exception exception, object state)
        {
            if (exception != null)
            {
                SocketException(this, new ErrorEventArgs(exception));
                return;
            }
            if (response == null)
            {
                ConnectionClosed(this, null);
            }
            List<String> responses;
            if (response.StartsWith("Version:"))
            {
                responses = (List<String>)state;
                responses.Add(response);
                _version = response;
                _socket.BeginReceive(XmlCallback, responses);
                return;
            }
            if (response.StartsWith("Length:"))
            {
                responses = (List<String>)state;
                responses.Add(response);
                _socket.BeginReceive(XmlCallback, responses);
                return;
            }


            // XML STUFF
            var name = string.Empty;
            try
            {
                using (var reader = XmlReader.Create(new StringReader(response)))
                {
                    while (reader.Read())
                    {
                        if (!reader.IsStartElement())
                        {
                            continue;
                        }
                        switch (reader.Name)
                        {
                            case "name": // Reads to name and grabs value.
                                reader.Read();
                                name = reader.Value;
                                break;
                            case "contents": // Reads to contents, grabs value and adds cell.
                                reader.Read();
                                var contents = reader.Value;
                                SetContentsOfCell(name, contents);
                                break;
                        }
                    }
                }
                Joined(this, null);
                _socket.BeginReceive(CallbackOperatorCallback, null);                
            }
            catch (Exception e)
            {
                throw new SpreadsheetReadWriteException("Error. " + e.Message);
            }
        }

        /// <summary>
        /// Handles the server response to the CREATE command and finally calls Join.
        /// </summary>
        /// <param name="response">The server response</param>
        /// <param name="exception">Any socket error</param>
        /// <param name="state">State object</param>
        private void CreateCallback(string response, Exception exception, object state)
        {
            if (exception != null)
            {
                SocketException(this, new ErrorEventArgs(exception));
                return;
            }
            if (response == null)
            {
                ConnectionClosed(this, null);
            }
            List<string> responses;
            if (exception != null)
            {
                // Raise error event here
                return;
            }
            if (response.Equals("CREATE FAIL"))
            {
                responses = new List<string>();
                responses.Add(response);
                _socket.BeginReceive(FailCallback, responses);
                return;
            }
            if (response.Equals("CREATE OK"))
            {
                responses = new List<string>();
                responses.Add(response);
                _socket.BeginReceive(CreateCallback, responses);
                return;
            }
            responses = (List<string>)state;
            responses.Add(response);
            if (response.StartsWith("Name:"))
            {
                _name = response;
                _socket.BeginReceive(CreateCallback, responses);
                return;
            }
            if (response.StartsWith("Password:"))
            {
                Join(_name.Substring(5), response.Substring(9));
                return;
            }
            _socket.BeginReceive(FailCallback, responses);
        }

        /// <summary>
        /// The master callback to be constantly listening after spreadsheet JOINed.
        /// Should be called again at the end of specific message handling callbacks.
        /// </summary>
        /// <param name="response">The server message</param>
        /// <param name="exception">Any socket exception</param>
        /// <param name="state">State object</param>
        private void CallbackOperatorCallback(string response, Exception exception, object state)
        {
            if (exception != null)
            {
                SocketException(this, new ErrorEventArgs(exception));
                return;
            }
            if (response == null)
            {
                ConnectionClosed(this, null);
                return;
            }
            if (response.StartsWith("UPDATE"))
            {
                UpdateCallback(response, exception, state);
            }
            if (response.StartsWith("CHANGE"))
            {
                ChangeCallback(response, exception, state);
            }
            if (response.StartsWith("UNDO"))
            {
                UndoCallback(response, exception, state);
            }
            if (response.StartsWith("SAVE"))
            {
                SaveCallback(response, exception, state);
            }
            if (response.Equals("ERROR"))
            {
                GenericError(this, null);
            }
        }

        /// <summary>
        /// Handles CHANGE response from server.
        /// </summary>
        /// <param name="response">The response</param>
        /// <param name="exception">Any exceptions</param>
        /// <param name="state">State object</param>
        private void ChangeCallback(string response, Exception exception, object state)
        {
            if (exception != null)
            {
                SocketException(this, new ErrorEventArgs(exception));
                return;
            }
            if (response == null)
            {
                ConnectionClosed(this, null);
            }
            List<string> responses;
            if (response.Equals("CHANGE FAIL"))
            {
                responses = new List<string>();
                responses.Add(response);
                _currentChangeCell = null;
                _currentChangeContent = null;
                _socket.BeginReceive(FailCallback, responses);
                return;
            }
            if (response.StartsWith("CHANGE"))
            {
                responses = new List<string>();
                responses.Add(response);
                _socket.BeginReceive(ChangeCallback, responses);
                return;
            }
            responses = (List<string>)state;
            responses.Add(response);
            if (response.StartsWith("Name:"))
            {
                _socket.BeginReceive(ChangeCallback, responses);
                return;
            }
            // We know that this will be Version:.
            if (responses.First().Equals("CHANGE OK"))
            {
                ISet<string> toBeRedisplayed = SetContentsOfCell(_currentChangeCell, _currentChangeContent);
                _currentChangeCell = null;
                _currentChangeContent = null;
                _version = response;
                Updated(this, new UpdatedEventArgs(toBeRedisplayed));
            }
            if (responses.First().Equals("CHANGE WAIT"))
            {
                if (response == _version)
                {
                    string currentChangeCell = _currentChangeCell;
                    string currentChangeContent = _currentChangeContent;
                    _currentChangeCell = null;
                    _currentChangeContent = null;
                    Change(currentChangeCell, currentChangeContent);
                }
                else
                {
                    _currentChangeCell = null;
                    _currentChangeContent = null;
                    //YouAreHosed(this, null);
                }
            }
            _socket.BeginReceive(CallbackOperatorCallback, null);
        }

        /// <summary>
        /// Handles UNDO responses from server.
        /// </summary>
        /// <param name="response">The response</param>
        /// <param name="exception">Any exceptions</param>
        /// <param name="state">State object</param>
        private void UndoCallback(string response, Exception exception, object state)
        {
            if (exception != null)
            {
                SocketException(this, new ErrorEventArgs(exception));
                return;
            }
            if (response == null)
            {
                ConnectionClosed(this, null);
            }
            List<string> responses;
            if (response.Equals("UNDO FAIL"))
            {
                responses = new List<string>();
                responses.Add(response);
                _socket.BeginReceive(FailCallback, responses);
                return;
            }
            if (response.StartsWith("UNDO"))
            {
                responses = new List<string>();
                responses.Add(response);
                _socket.BeginReceive(UndoCallback, responses);
                return;
            }
            responses = (List<string>)state;
            responses.Add(response);
            if (response.StartsWith("Name:"))
            {
                _socket.BeginReceive(UndoCallback, responses);
                return;
            }
            if (response.StartsWith("Version:"))
            {
                // If the first response was UNDO END or UNDO WAIT, then we're
                // at the last part of the message. Otherwise, we aren't.
                if (responses.First().Equals("UNDO END"))
                {
                    UndoEnd(this, null);
                    _socket.BeginReceive(CallbackOperatorCallback, null);
                    return;
                }
                if (responses.First().Equals("UNDO WAIT"))
                {
                    if (response.Equals(_version))
                    {
                        Undo();
                    }
                    else
                    {
                        //YouAreHosed(this, null);
                    }
                    _socket.BeginReceive(CallbackOperatorCallback, null);
                    return;
                }
                _socket.BeginReceive(UndoCallback, responses);
                return;
            }
            if (response.StartsWith("Cell:") || response.StartsWith("Length:"))
            {
                _socket.BeginReceive(UndoCallback, responses);
                return;
            }
            ISet<string> toBeRedisplayed = SetContentsOfCell(responses.ElementAt(3).Substring(5), response);
            _version = responses.ElementAt(2);
            Updated(this, new UpdatedEventArgs(toBeRedisplayed));
            _socket.BeginReceive(CallbackOperatorCallback, null);
        }

        /// <summary>
        /// Handles SAVE responses from server.
        /// </summary>
        /// <param name="response">The response</param>
        /// <param name="exception">Any exceptions</param>
        /// <param name="state">State object</param>
        private void SaveCallback(string response, Exception exception, object state)
        {
            if (exception != null)
            {
                SocketException(this, new ErrorEventArgs(exception));
                return;
            }
            if (response == null)
            {
                ConnectionClosed(this, null);
            }
            if (response.Equals("SAVE FAIL"))
            {
                List<string> responses = new List<string>();
                responses.Add(response);
                _socket.BeginReceive(FailCallback, responses);
                return;
            }
            if (response.StartsWith("SAVE"))
            {
                _socket.BeginReceive(SaveCallback, null);
                return;
            }
            _socket.BeginReceive(CallbackOperatorCallback, null);
        }

        /// <summary>
        /// Send CREATE command to server
        /// </summary>
        /// <param name="filename">Name of the file to create</param>
        /// <param name="password">Password for the file to create</param>
        public void Create(string filename, string password)
        {
            _socket.BeginSend("CREATE\nName:" + filename + "\nPassword:" + password + "\n", (e, o) => { }, null);
            _socket.BeginReceive(CreateCallback, null);
        }

        /// <summary>
        /// Verifies that no circular exception will be created and then sends CHANGE request to server.
        /// </summary>
        /// <param name="name">The cell to be changed</param>
        /// <param name="content">The candidate contents of the cell</param>
        public void Change(string name, string content)
        {
            // If we're waiting for a change to be confirmed by the server, don't take a new change.
            if (_currentChangeCell != null)
            {
                return;
            }
            var normalizedName = Normalize(name);

            // Check if content is null.
            if (content == null)
            {
                throw new ArgumentNullException();
            }
            // Check if name is null or invalid.
            if (normalizedName == null || !_validCellNameRegex.IsMatch(normalizedName) || !IsValid(normalizedName))
            {
                throw new InvalidNameException();
            }
            if (content.StartsWith("="))
            {
                Formula formula = new Formula(content.Substring(1), IsValid, Normalize);
            
                // The HashSet dependents contains name and all of name's direct and indirect dependents.
                var dependents = new HashSet<string>(GetCellsToRecalculate(normalizedName));

                // Variables contains name's new dependees.
                var variables = formula.GetVariables();

                // Checking if any of name's new dependees are already its dependents
                // or if name is its own new dependee.
                if (dependents.Overlaps(variables))
                {
                    throw new CircularException();
                }
            }
            _currentChangeCell = normalizedName;
            _currentChangeContent = content;
            _socket.BeginSend("CHANGE\n" + _name + "\n" + _version + "\nCell:" + normalizedName + "\nLength:" + content.Length.ToString() + "\n" + content + "\n", (e, o) => { }, null);
        }

        /// <summary>
        /// Sends the undo request to the server.
        /// </summary>
        public void Undo()
        {
            _socket.BeginSend("UNDO\n" + _name + "\n" + _version + "\n", (e, o) => { }, null);
        }

        /// <summary>
        /// Sends the save request to the server.
        /// </summary>
        public void Save()
        {
            _socket.BeginSend("SAVE\n" + _name + "\n", (e, o) => { }, null);
        }

        /// <summary>
        /// Tells the server that we're leaving.
        /// </summary>
        public void Leave()
        {
            if (_socket == null) return;
            _socket.BeginSend("LEAVE\n" + _name + "\n", (e, o) => { }, null);
            _socket.Close();
        }

        /// <summary>
        /// Handles UPDATE messages from server.
        /// </summary>
        /// <param name="response">The server message</param>
        /// <param name="exception">Any socket exception</param>
        /// <param name="state">State object</param>
        private void UpdateCallback(string response, Exception exception, object state)
        {
            if (exception != null)
            {
                SocketException(this, new ErrorEventArgs(exception));
                return;
            }
            if (response == null)
            {
                ConnectionClosed(this, null);
            }
            List<string> responses;
            if (response.Equals("UPDATE"))
            {
                responses = new List<string>();
                responses.Add(response);
                _socket.BeginReceive(UpdateCallback, responses);
                return;
            }
            responses = (List<string>)state;
            responses.Add(response);
            if (response.StartsWith("Name:"))
            {
                _socket.BeginReceive(UpdateCallback, responses);
                return;
            }
            if(response.StartsWith("Version:"))
            {
                _version = response;
                _socket.BeginReceive(UpdateCallback, responses);
                return;
            }
            if (response.StartsWith("Cell:"))
            {
                _socket.BeginReceive(UpdateCallback, responses);
                return;
            }
            if (response.StartsWith("Length:"))
            {
                _socket.BeginReceive(UpdateCallback, responses);
                return;
            }
            ISet<string> toBeRedisplayed = SetContentsOfCell(responses.ElementAt(3).Substring(5), response);
            Updated(this, new UpdatedEventArgs(toBeRedisplayed));
            _socket.BeginReceive(CallbackOperatorCallback, null);
        }


        /// <summary>
        ///     Enumerates the names of all non-empty cells in the spreadsheet.
        /// </summary>
        public override IEnumerable<string> GetNamesOfAllNonemptyCells()
        {
            return _cells.Keys;
        }

        /// <summary>
        ///     If name is null or invalid, throws an InvalidNameException.
        ///     Otherwise, returns the contents (as opposed to the value) of the named cell.  The return
        ///     value should be either a string, a double, or a Formula.
        /// </summary>
        public override object GetCellContents(string name)
        {
            Cell cell; // Unitialized as the out parameter will put a value there.

            // Normalizing before checking validity.
            var normalizedName = Normalize(name);
            // Checking for a valid name.
            if (normalizedName == null || !_validCellNameRegex.IsMatch(normalizedName) || !IsValid(normalizedName))
            {
                throw new InvalidNameException();
            }
            // Check if there is content in the named cell.
            return _cells.TryGetValue(normalizedName, out cell) ? cell.Contents : String.Empty;
            // An empty cell has an empty string as its contents.
        }

        /// <summary>
        ///     If name is null or invalid, throws an InvalidNameException.
        ///     Otherwise, returns the value (as opposed to the contents) of the named cell.  The return
        ///     value should be either a string, a double, or a SpreadsheetUtilities.FormulaError.
        /// </summary>
        public override object GetCellValue(string name)
        {
            Cell cell;
            // Normalizing name.
            var normalizeName = Normalize(name);
            // Check if name is a valid cell name.
            if (normalizeName == null || !_validCellNameRegex.IsMatch(normalizeName) || !IsValid(normalizeName))
            {
                throw new InvalidNameException();
            }
            // Check if named cell has a value.
            return _cells.TryGetValue(normalizeName, out cell) ? cell.Value : string.Empty;
            // Return the value of an empty string, if name is empty
        }

        /// <summary>
        ///     If content is null, throws an ArgumentNullException.
        ///     Otherwise, if name is null or invalid, throws an InvalidNameException.
        ///     Otherwise, if content parses as a double, the contents of the named
        ///     cell becomes that double.
        ///     Otherwise, if content begins with the character '=', an attempt is made
        ///     to parse the remainder of content into a Formula f using the Formula
        ///     constructor.  There are then three possibilities:
        ///     (1) If the remainder of content cannot be parsed into a Formula, a
        ///     SpreadsheetUtilities.FormulaFormatException is thrown.
        ///     (2) Otherwise, if changing the contents of the named cell to be f
        ///     would cause a circular dependency, a CircularException is thrown.
        ///     (3) Otherwise, the contents of the named cell becomes f.
        ///     Otherwise, the contents of the named cell becomes content.
        ///     If an exception is not thrown, the method returns a set consisting of
        ///     name plus the names of all other cells whose value depends, directly
        ///     or indirectly, on the named cell.
        ///     For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        ///     set {A1, B1, C1} is returned.
        /// </summary>
        public override sealed ISet<string> SetContentsOfCell(string name, string content)
        {
            var normalizedName = Normalize(name);
            double numberContents;
            // Check if content is null.
            if (content == null)
            {
                throw new ArgumentNullException();
            }
            // Check if name is null or invalid.
            if (normalizedName == null || !_validCellNameRegex.IsMatch(normalizedName) || !IsValid(normalizedName))
            {
                throw new InvalidNameException();
            }
            // Try setting cell contents to double.
            if (double.TryParse(content, out numberContents))
            {
                return SetCellContents(name, numberContents);
            }
            // Try setting cell contents to formula.
            return content.StartsWith("=") ? SetCellContents(name, new Formula(content.Substring(1), IsValid, Normalize)) : SetCellContents(name, content);
            // Or else set cell contents to string.
        }

        /// <summary>
        ///     If name is null or invalid, throws an InvalidNameException.
        ///     Otherwise, the contents of the named cell becomes number.  The method returns a
        ///     set consisting of name plus the names of all other cells whose value depends,
        ///     directly or indirectly, on the named cell.
        ///     For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        ///     set {A1, B1, C1} is returned.
        /// </summary>
        protected override ISet<string> SetCellContents(string name, double number)
        {
            // Normalizing name.
            var normalizedName = Normalize(name);
            if (normalizedName == null || !_validCellNameRegex.IsMatch(normalizedName) || !IsValid(normalizedName))
            {
                throw new InvalidNameException();
            }
            // Name is no longer a dependent of any other cell.
            _dependencies.ReplaceDependees(normalizedName, new HashSet<string>());
            // Setting contents of cell
            _cells[normalizedName] = new Cell(number);

            var cellsToRecalculate = GetCellsToRecalculate(normalizedName);
            // Updating values (in order) of each affected cell.
            var localCellsToRecalculate = cellsToRecalculate as string[] ?? cellsToRecalculate.ToArray();
            RecalculateCellValues(localCellsToRecalculate);

            Changed = true;

            // Getting set of cells whose value depends on name.
            return cellsToRecalculate != null ? new HashSet<string>(localCellsToRecalculate) : new HashSet<string>();
        }

        /// <summary>
        ///     If text is null, throws an ArgumentNullException.
        ///     Otherwise, if name is null or invalid, throws an InvalidNameException.
        ///     Otherwise, the contents of the named cell becomes text.  The method returns a
        ///     set consisting of name plus the names of all other cells whose value depends,
        ///     directly or indirectly, on the named cell.
        ///     For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        ///     set {A1, B1, C1} is returned.
        /// </summary>
        protected override ISet<string> SetCellContents(string name, string text)
        {
            // Normalizing name.
            var normalizedName = Normalize(name);
            if (text == null)
            {
                throw new ArgumentNullException();
            }
            if (normalizedName == null || !_validCellNameRegex.IsMatch(normalizedName) || !IsValid(normalizedName))
            {
                throw new InvalidNameException();
            }
            // Text cells don't depend on other cells, so eliminate their dependees.
            _dependencies.ReplaceDependees(normalizedName, new HashSet<string>());

            if (text != string.Empty)
            {
                _cells[normalizedName] = new Cell(text);
            }

                // If a cell has an empty string as its contents, it has an empty string
                // as its value and it is itself an empty cell in which case we don't
                // need to keep track of it any longer.
            else
            {
                _cells.Remove(normalizedName);
            }

            var cellsToRecalculate = GetCellsToRecalculate(normalizedName);
            var localCellsToRecalculate = cellsToRecalculate as string[] ?? cellsToRecalculate.ToArray();
            
            // Updating values (in order) of each affected cell.
            RecalculateCellValues(localCellsToRecalculate);

            Changed = true;
            // Returning set of cells whose values depend on name.
            return new HashSet<string>(localCellsToRecalculate);
        }

        /// <summary>
        ///     If formula parameter is null, throws an ArgumentNullException.
        ///     Otherwise, if name is null or invalid, throws an InvalidNameException.
        ///     Otherwise, if changing the contents of the named cell to be the formula would cause a
        ///     circular dependency, throws a CircularException.
        ///     Otherwise, the contents of the named cell becomes formula.  The method returns a
        ///     Set consisting of name plus the names of all other cells whose value depends,
        ///     directly or indirectly, on the named cell.
        ///     For example, if name is A1, B1 contains A1*2, and C1 contains B1+A1, the
        ///     set {A1, B1, C1} is returned.
        /// </summary>
        protected override ISet<string> SetCellContents(string name, Formula formula)
        {
            // Normaling name.
            var normalizedName = Normalize(name);
            if (formula == null)
            {
                throw new ArgumentNullException();
            }
            if (normalizedName == null || !_validCellNameRegex.IsMatch(normalizedName) || !IsValid(normalizedName))
            {
                throw new InvalidNameException();
            }

            var cellsToRecalculate = GetCellsToRecalculate(normalizedName);
            var localCellsToRecalculate = cellsToRecalculate as string[] ?? cellsToRecalculate.ToArray();


            _cells[normalizedName] = new Cell(formula);

            // Replacing name's dependees.
            _dependencies.ReplaceDependees(normalizedName, formula.GetVariables());

            // Updating values (in order) of each affected cell.
            RecalculateCellValues(localCellsToRecalculate);

            Changed = true;
            return new HashSet<string>(localCellsToRecalculate);
        }

        /// <summary>
        ///     A helper for the GetCellsToRecalculate method.
        /// </summary>
        protected override IEnumerable<string> GetDirectDependents(string name)
        {
            var normalizedName = Normalize(name);
            if (normalizedName == null)
            {
                throw new ArgumentNullException();
            }
            if (!_validCellNameRegex.IsMatch(normalizedName) || !IsValid(normalizedName))
            {
                throw new InvalidNameException();
            }
            return _dependencies.GetDependents(normalizedName);
        }

        /// <summary>
        ///     A helper for SetCellContents which recalculates and sets the value of a cell.
        /// </summary>
        /// <param name="names">Cell names to recalculate.</param>
        private void RecalculateCellValues(IEnumerable<string> names)
        {
            // Checks if a cell has content.
            foreach (var name in names)
            {
                Cell cell;
                if (!_cells.TryGetValue(name, out cell))
                {
                    continue;
                }
                // Checks if the cell contains a formula. If so, evaluates the formula and
                // places formula's value in the cell's value field.
                var formula = cell.Contents as Formula;
                cell.Value = formula != null ? formula.Evaluate(Lookup) : cell.Contents;
            }
            // If the cell is empty, then it's value doesn't need to be tracked, as we know its
            // value (and content) is an empty string.
        }

        private double Lookup(string name)
        {
            // Retrieves value from cell.
            var value = GetCellValue(name);
            // Checks if it is a double, and if so, returns the value.
            if (value is double)
            {
                return (double) value;
            }
            // If not a double, throws an ArgumentException.
            throw new ArgumentException("Unknown variable.");
        }

        /// <summary>
        ///     Writes the contents of this spreadsheet to the named file using an XML format.
        ///     The XML elements should be structured as follows:
        ///     <spreadsheet version="version information goes here">
        ///         <cell>
        ///             <name>
        ///                 cell name goes here
        ///             </name>
        ///             <contents>
        ///                 cell contents goes here
        ///             </contents>
        ///         </cell>
        ///     </spreadsheet>
        ///     There should be one cell element for each non-empty cell in the spreadsheet.
        ///     If the cell contains a string, it should be written as the contents.
        ///     If the cell contains a double d, d.ToString() should be written as the contents.
        ///     If the cell contains a Formula f, f.ToString() with "=" prepended should be written as the contents.
        ///     If there are any problems opening, writing, or closing the file, the method should throw a
        ///     SpreadsheetReadWriteException with an explanatory message.
        /// </summary>
        public override void Save(string filename)
        {
            try
            {
                using (var writer = XmlWriter.Create(filename))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("spreadsheet");
                    writer.WriteAttributeString("version", Version);

                    foreach (var pair in _cells)
                    {
                        writer.WriteStartElement("cell");
                        writer.WriteStartElement("name");
                        writer.WriteString(pair.Key);
                        writer.WriteEndElement();
                        writer.WriteStartElement("contents");
                        // Adding an equals for formulas means that I can feed every read cell
                        // back into SetContentsOfCell and let it rebuild my spreadsheet.
                        var formula = pair.Value.Contents as Formula;
                        if (formula != null)
                        {
                            writer.WriteString("=" + formula);
                        }
                        else
                        {
                            writer.WriteString(pair.Value.Contents.ToString());
                        }
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
            }
            catch (Exception e)
            {
                throw new SpreadsheetReadWriteException("Error. " + e.Message);
            }
            Changed = false;
        }

        /// <summary>
        ///     Returns the version information of the spreadsheet saved in the named file.
        ///     If there are any problems opening, reading, or closing the file, the method
        ///     should throw a SpreadsheetReadWriteException with an explanatory message.
        /// </summary>
        public override sealed string GetSavedVersion(string filename)
        {
            try
            {
                using (var reader = XmlReader.Create(filename))
                {
                    while (reader.Read())
                    {
                        if (!reader.IsStartElement())
                        {
                            continue;
                        }
                        switch (reader.Name)
                        {
                            case "spreadsheet":
                                return reader["version"];
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new SpreadsheetReadWriteException("Error. " + e.Message);
            }
            throw new SpreadsheetReadWriteException("Unknown error.");
        }

        /// <summary>
        ///     Represents a spreadsheet cell where Contents represent the contents of the cell.
        /// </summary>
        private class Cell
        {
            // Representation Invariants:
            // Contents must be either a string, a double, or a Formula, and
            // may not be null or an empty string.
            // Value may either a string, a double, or a FormulaError, and may not be null.

            // Abstraction function:
            // The string, double, or Formula of Contents is the contents of the cell.
            // The string, double or FormulaError in Value is the value of the cell.
            public readonly object Contents;
            public object Value;

            /// <summary>
            ///     Constructs a cell.
            /// </summary>
            public Cell(object contents)
            {
                Contents = contents;
                Value = string.Empty;
            }
        }
    }
}