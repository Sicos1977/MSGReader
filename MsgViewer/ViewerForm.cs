using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using DocumentServices.Modules.Readers.MsgReader;

namespace MsgViewer
{
    public partial class ViewerForm : Form
    {
        List<string> _tempFolders = new List<string>(); 

        public ViewerForm()
        {
            InitializeComponent();
        }

        private void ViewerForm_Load(object sender, EventArgs e)
        {
            this.Closed += new EventHandler(ViewerForm_Closed);
        }

        void ViewerForm_Closed(object sender, EventArgs e)
        {
            foreach (var tempFolder in _tempFolders)
            {
                if (Directory.Exists(tempFolder))
                    Directory.Delete(tempFolder, true);
            }
        }

        private void SelectButton_Click(object sender, EventArgs e)
        {
            // Create an instance of the open file dialog box.
            var openFileDialog1 = new OpenFileDialog
            {
                Filter = "MSG Files (.msg)|*.msg",
                FilterIndex = 1,
                Multiselect = false
            };

            // Process input if the user clicked OK.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Open the selected file to read.
                string tempFolder = null;

                try
                {
                    tempFolder = GetTemporaryFolder();
                    _tempFolders.Add(tempFolder);
                    var msgReader = new Reader();
                    var files = msgReader.ExtractToFolder(openFileDialog1.FileName, tempFolder);

                    webBrowser1.Navigate(files[0]);
                }
                catch (Exception)
                {
                    if (tempFolder != null && Directory.Exists(tempFolder))
                        Directory.Delete(tempFolder, true);
                }
            }
        }

        public string GetTemporaryFolder()
        {
            var tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }
    }
}
