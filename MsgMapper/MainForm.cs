using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using DocumentServices.Modules.Readers.MsgReader;
using DocumentServices.Modules.Readers.MsgReader.Outlook;
using MsgMapper.Properties;

namespace MsgMapper
{
    public partial class MainForm : Form
    {
        // This delegate enables asynchronous calls for setting
        // the text property on a TextBox control.
        delegate void SetListBoxCallback(string text);

        #region Fields
        /// <summary>
        /// Contains our notifyicon on the taskbar
        /// </summary>
        private readonly NotifyIcon _trayIcon;

        /// <summary>
        /// A list with one or more <see cref="FileSystemWatcher">FileSystemWatchers</see>
        /// </summary>
        private readonly List<FileSystemWatcher> _fileSystemWatcher = new List<FileSystemWatcher>();

        /// <summary>
        /// Flag used to signal the <see cref="MainForm_FormClosing"/> method to really exit
        /// </summary>
        private bool _exit;

        /// <summary>
        /// The <see cref="Reader"/> object that does all our properties mapping
        /// </summary>
        private Reader _reader = new Reader();
        #endregion

        #region MainForm
        public MainForm()
        {
            InitializeComponent();

            var trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Show", OnShow);
            trayMenu.MenuItems.Add("Resync", OnResync);
            trayMenu.MenuItems.Add("Exit", OnExit);

            // Create a tray icon
            _trayIcon = new NotifyIcon
            {
                // ReSharper disable once LocalizableElement
                Text = "MSG properties to extended file properties mapper",
                Icon = new Icon(Icon, 40, 40),
                ContextMenu = trayMenu,
                Visible = true
            };
        }
        #endregion

        #region _fileSystemWatcher_Created
        /// <summary>
        /// This method is called when a new <see cref="Storage.Message"/> object is created in a watched folder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void _fileSystemWatcher_Created(object sender, FileSystemEventArgs e)
        {
            var msgFile = string.Empty;

            try
            {
                msgFile = e.FullPath;
                _reader.SetExtendedFileAttributesWithMsgProperties(msgFile);
                Log("Mapped properties for the file '" + msgFile + "'");
            }
            catch (Exception exception)
            {
                Log("Failed to map properties for the file '" + msgFile + "', error: " + exception.Message);
            }
        }
        #endregion

        #region OnShow
        /// <summary>
        /// This method is triggered when a user selects the <see cref="ContextMenu"/> item called Show
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnShow(object sender, EventArgs e)
        {
            Show();
        }
        #endregion

        #region OnResync
        /// <summary>
        /// This method is triggered when a user selects the <see cref="ContextMenu"/> item called Resync
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnResync(object sender, EventArgs e)
        {
            if (
                MessageBox.Show(
                    // ReSharper disable LocalizableElement
                    "Are you sure you want to resync all the folers? This can take some time when you have huge amounts of Outlook message files",
                    "Resync folders", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    // ReSharper restore LocalizableElement
            {
                foreach (DataGridViewRow row in FoldersDataGridView.Rows)
                {
                    if (row.Cells[0].Value == null) continue;
                    var folder = row.Cells[0].Value.ToString();
                    Log("Resyncing folder '" + folder + "'");
                    var msgFiles = Directory.GetFiles(folder, "*.msg", SearchOption.AllDirectories);
                    foreach (var msgFile in msgFiles)
                    {
                        _reader.SetExtendedFileAttributesWithMsgProperties(msgFile);
                        Log("Mapped properties for the file '" + msgFile + "'");
                    }
                }
                Log("Done");
            }
        }
        #endregion

        #region OnExit
        /// <summary>
        /// This method is triggered when a user selects the <see cref="ContextMenu"/> item called Exit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnExit(object sender, EventArgs e)
        {
            SaveFolders();
            _exit = true;
            Application.Exit();
        }
        #endregion

        #region OnLoad
        /// <summary>
        /// This method is triggered when the form loads for the first time
        /// </summary>
        /// <param name="e"></param>
        protected override void OnLoad(EventArgs e)
        {
            Visible       = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.
            LoadFolders();
            base.OnLoad(e);
        }
        #endregion

        #region Log
        /// <summary>
        /// Writes messages to the <see cref="StatusListBox"/>
        /// </summary>
        /// <param name="message"></param>
        private void Log(string message)
        {
            // InvokeRequired required compares the thread ID of the 
            // calling thread to the thread ID of the creating thread. 
            // If these threads are different, it returns true. 
            if (InvokeRequired)
            {
                var callback = new SetListBoxCallback(Log);
                Invoke(callback, new object[] {message});
            }
            else
            {
                StatusListBox.Items.Add(message);
                StatusListBox.SelectedIndex = StatusListBox.Items.Count - 1;
            }
        }
        #endregion

        #region LoadFolders
        /// <summary>
        /// This will load all the selected folders and add them to the filesystem watcher
        /// </summary>
        private void LoadFolders()
        {
            var foldersString = (string) Settings.Default["Folders"];
            try
            {
                if (string.IsNullOrEmpty(foldersString))
                {
                    Show();
                }
                else
                {
                    var folders = Folders.LoadFromString(foldersString);
                    foreach (var folder in folders)
                    {
                        // ReSharper disable once PossiblyMistakenUseOfParamsMethod
                        if (Directory.Exists(folder.Path))
                        {
                            FoldersDataGridView.Rows.Add(folder.Path);
                            var fileSystemWatcher = new FileSystemWatcher();
                            fileSystemWatcher.Path = folder.Path;
                            fileSystemWatcher.IncludeSubdirectories = true;
                            fileSystemWatcher.Filter = "*.msg";
                            fileSystemWatcher.Created += _fileSystemWatcher_Created;
                            fileSystemWatcher.EnableRaisingEvents = true;
                            _fileSystemWatcher.Add(fileSystemWatcher);
                            Log("Watching folder '" + folder.Path + "'");
                        }
                        else
                        {
                            Log("Ignoring folder '" + folder.Path + "' because it does not exists anymore");
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region SaveFolders
        /// <summary>
        /// This will save all the selected folders to the <see cref="Settings"/> object
        /// </summary>
        private void SaveFolders()
        {
            var folders = new Folders();

            foreach (DataGridViewRow row in FoldersDataGridView.Rows)
            {
                if(row.Cells[0].Value == null) continue;
                folders.Add(row.Cells[0].Value.ToString());
            }

            // Save setting
            Settings.Default["Folders"] = folders.SerializeToString();
            Settings.Default.Save();
        }
        #endregion

        #region MainForm_FormClosing
        /// <summary>
        /// This method is triggered when a user tries to close the application. When a users presses the X (top right)
        /// on the form the form is just hidden. When the method is triggered through the <see cref="OnExit"/> method the
        /// form will be closed and the application with exit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_exit) return;

            // Do not close the form, just hide it... except when we close from the OnExit method
            Hide();
            e.Cancel = true;
        }
        #endregion

        #region FolderExists
        /// <summary>
        /// This method checks if the <see cref="folder"/> is already present in the <see cref="FoldersDataGridView"/>
        /// </summary>
        /// <param name="folder"></param>
        /// <returns></returns>
        private bool FolderExists(string folder)
        {
            foreach (var row in FoldersDataGridView.Rows)
                if (String.Equals(row.ToString(), folder, StringComparison.InvariantCultureIgnoreCase))
                    return true;

            return false;
        }
        #endregion

        #region AddFolderButton_Click
        /// <summary>
        /// This method adds a folder to the <see cref="FoldersDataGridView"/>
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddFolderButton_Click(object sender, EventArgs e)
        {
            var folderBrowserDialog = new FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                // ReSharper disable once PossiblyMistakenUseOfParamsMethod
                if (FolderExists(folderBrowserDialog.SelectedPath))
                    MessageBox.Show("The folder '" + folderBrowserDialog.SelectedPath + "' is already added");
                else
                    FoldersDataGridView.Rows.Add(folderBrowserDialog.SelectedPath);
            }
        }
        #endregion
    }
}
