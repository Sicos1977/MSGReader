using MsgReader;
using MsgViewer.Helpers;
using MsgViewer.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;

//
// Author: Kees van Spelde <sicos2002@hotmail.com>
//
// Copyright (c) 2013-2024 Kees van Spelde. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NON INFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

namespace MsgViewer
{
    public partial class ViewerForm : Form
    {
        #region Fields
        /// <summary>
        /// Used to track all the created temporary folders
        /// </summary>
        readonly List<string> _tempFolders = [];
        #endregion

        #region Form events
        public ViewerForm()
        {
            InitializeComponent();
        }

        private void ViewerForm_Load(object sender, EventArgs e)
        {
            WindowPlacement.SetPlacement(Handle, Settings.Default.Placement);
            Closing += ViewerForm_Closing;
            generateHyperlinksToolStripMenuItem.Checked = Settings.Default.GenerateHyperLinks;
            //var version = Assembly.GetExecutingAssembly().GetName().Version;

            // ReSharper disable LocalizableElement
            Text = "MSG Viewer v5.6.1";
            // ReSharper restore LocalizableElement

            SetCulture(Settings.Default.Language);

            var args = Environment.GetCommandLineArgs();

            if (args.Length > 1 && File.Exists(args[1]))
                OpenFile(args[1]);
        }

        private void ViewerForm_Closing(object sender, EventArgs e)
        {
            Settings.Default.Placement = WindowPlacement.GetPlacement(Handle);
            Settings.Default.Save();
            foreach (var tempFolder in _tempFolders)
            {
                if (Directory.Exists(tempFolder))
                    Directory.Delete(tempFolder, true);
            }
        }
        #endregion

        #region GetTemporaryFolder
        /// <summary>
        /// Returns a temporary folder
        /// </summary>
        /// <returns></returns>
        private static string GetTemporaryFolder()
        {
            var tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }
        #endregion

        #region WebBrowser events
        private void BackButton_Click_1(object sender, EventArgs e)
        {
            webBrowser1.GoBack();
        }

        private void ForwardButton_Click_1(object sender, EventArgs e)
        {
            webBrowser1.GoForward();
        }

        private void PrintButton_Click(object sender, EventArgs e)
        {
            webBrowser1.ShowPrintDialog();
        }

        private void webBrowser1_Navigated_1(object sender, WebBrowserNavigatedEventArgs e)
        {
            if (e.Url != null) StatusLabel.Text = e.Url.ToString();
            BackButton.Enabled = webBrowser1.CanGoBack;
            ForwardButton.Enabled = webBrowser1.CanGoForward;
        }
        #endregion

        #region SaveAsTextButton_Click
        private void SaveAsTextButton_Click(object sender, EventArgs e)
        {
            // Create an instance of the save file dialog box.
            var saveFileDialog1 = new SaveFileDialog
            {
                // ReSharper disable once LocalizableElement
                Filter = "TXT Files (.txt)|*.txt",
                FilterIndex = 1
            };

            if (Directory.Exists(Settings.Default.SaveDirectory))
                saveFileDialog1.InitialDirectory = Settings.Default.SaveDirectory;

            // Process input if the user clicked OK.
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Settings.Default.SaveDirectory = Path.GetDirectoryName(saveFileDialog1.FileName);
                var htmlToText = new HtmlToText();
                var text = htmlToText.Convert(webBrowser1.DocumentText);
                File.WriteAllText(saveFileDialog1.FileName, text);
            }
        }
        #endregion

        #region openToolStripMenuItem_Click
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Create an instance of the open file dialog box.
            var openFileDialog1 = new OpenFileDialog
            {
                // ReSharper disable once LocalizableElement
                Filter = "E-mail|*.msg;*.eml",
                FilterIndex = 1,
                Multiselect = false
            };

            if (Directory.Exists(Settings.Default.InitialDirectory))
                openFileDialog1.InitialDirectory = Settings.Default.InitialDirectory;

            // Process input if the user clicked OK.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Settings.Default.InitialDirectory = Path.GetDirectoryName(openFileDialog1.FileName);
                OpenFile(openFileDialog1.FileName);
            }
        }
        #endregion

        #region printToolStripMenuItem_Click
        private void printToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.ShowPrintDialog();
        }
        #endregion

        #region exitToolStripMenuItem_Click
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }
        #endregion

        #region OpenFile
        /// <summary>
        /// Opens the selected MSG of EML file
        /// </summary>
        /// <param name="fileName"></param>
        private void OpenFile(string fileName)
        {
            // Open the selected file to read.
            string tempFolder = null;

            try
            {
                tempFolder = GetTemporaryFolder();
                _tempFolders.Add(tempFolder);

                var msgReader = new Reader();

                // Use this, if you want to extract the code in memory
                // using (var streamReader = new StreamReader(fileName))
                // {
                //     string _body = msgReader.ExtractMsgEmailBody(streamReader.BaseStream, true, "text/html; charset=utf-8");
                // }

                var files = msgReader.ExtractToFolder(
                    fileName,
                    tempFolder,
                    generateHyperlinksToolStripMenuItem.Checked ? ReaderHyperLinks.Both : ReaderHyperLinks.None);

                // Use this, if you want to display a header table elsewhere but not in the web browser
                // var header = msgReader.ExtractMsgEmailHeader(new Storage.Message(fileName), true);

                var error = msgReader.GetErrorMessage();

                if (!string.IsNullOrEmpty(error))
                    throw new Exception(error);

                if (!string.IsNullOrEmpty(files[0]))
                    webBrowser1.Navigate(files[0]);

                FilesListBox.Items.Clear();

                foreach (var file in files)
                {
                    FilesListBox.Items.Add(file);
                }
            }
            catch (Exception ex)
            {
                if (tempFolder != null && Directory.Exists(tempFolder))
                    Directory.Delete(tempFolder, true);

                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region GenerateHyperlinksToolStripMenuItem_Click
        private void GenerateHyperlinksToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (generateHyperlinksToolStripMenuItem.Checked)
            {
                generateHyperlinksToolStripMenuItem.Checked = true;
                Settings.Default.GenerateHyperLinks = true;
            }
            else
            {
                generateHyperlinksToolStripMenuItem.Checked = false;
                Settings.Default.GenerateHyperLinks = false;
            }
            Settings.Default.Save();
        }
        #endregion

        #region LanguageToolStripMenuItem_Click
        private void LanguageToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sender == LanguageEnglishMenuItem)
                Settings.Default.Language = 1;
            else if (sender == LanguageFrenchMenuItem)
                Settings.Default.Language = 2;
            else if (sender == LanguageGermanMenuItem)
                Settings.Default.Language = 3;
            else if (sender == LanguageDutchMenuItem)
                Settings.Default.Language = 4;
            else if (sender == LanguageSpanishMenuItem)
                Settings.Default.Language = 5;
            else if (sender == LanguageSimpChineseMenuItem)
                Settings.Default.Language = 6;
            else if (sender == LanguageTradChineseMenuItem)
                Settings.Default.Language = 7;

            SetCulture(Settings.Default.Language);
            Settings.Default.Save();
        }
        #endregion

        #region SetCulture
        /// <summary>
        /// Sets the culture
        /// </summary>
        /// <param name="culture"></param>
        private void SetCulture(int culture)
        {
            LanguageEnglishMenuItem.Checked = false;
            LanguageFrenchMenuItem.Checked = false;
            LanguageGermanMenuItem.Checked = false;
            LanguageDutchMenuItem.Checked = false;
            LanguageSpanishMenuItem.Checked = false;
            LanguageSimpChineseMenuItem.Checked = false;
            LanguageTradChineseMenuItem.Checked = false;

            switch (culture)
            {
                case 1:
                    Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-US");
                    LanguageEnglishMenuItem.Checked = true;
                    break;
                case 2:
                    Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo("fr-FR");
                    LanguageFrenchMenuItem.Checked = true;
                    break;
                case 3:
                    Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo("de-DE");
                    LanguageGermanMenuItem.Checked = true;
                    break;
                case 4:
                    Thread.CurrentThread.CurrentCulture = new CultureInfo("nl-NL");
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo("nl-NL");
                    LanguageDutchMenuItem.Checked = true;
                    break;
                case 5:
                    Thread.CurrentThread.CurrentCulture = new CultureInfo("es-ES");
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo("es-ES");
                    LanguageSpanishMenuItem.Checked = true;
                    break;
                case 6:
                    Thread.CurrentThread.CurrentCulture = new CultureInfo("zh-CN");
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo("zh-CN");
                    LanguageSimpChineseMenuItem.Checked = true;
                    break;
                case 7:
                    Thread.CurrentThread.CurrentCulture = new CultureInfo("zh-TW");
                    Thread.CurrentThread.CurrentUICulture = new CultureInfo("zh-TW");
                    LanguageTradChineseMenuItem.Checked = true;
                    break;
            }
        }
        #endregion

        #region FilesListBox_DoubleClick
        private void FilesListBox_DoubleClick(object sender, EventArgs e)
        {
            if (FilesListBox.Items.Count <= 0) return;
            var file = FilesListBox.SelectedItem as string;
            if (string.IsNullOrEmpty(file) || !File.Exists(file)) return;
            var fileInfo = new FileInfo(file);
            if (fileInfo.Extension.ToLowerInvariant() == ".msg")
                Process.Start(Application.ExecutablePath, file);
            else
                Process.Start(file);
        }
        #endregion
    }
}
