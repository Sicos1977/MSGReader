namespace MsgViewer
{
    partial class ViewerForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ViewerForm));
            Properties.Settings settings1 = new Properties.Settings();
            statusStrip1 = new System.Windows.Forms.StatusStrip();
            StatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            panel2 = new System.Windows.Forms.Panel();
            FilesListBox = new System.Windows.Forms.ListBox();
            label1 = new System.Windows.Forms.Label();
            menuStrip1 = new System.Windows.Forms.MenuStrip();
            fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            printToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            languageToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            LanguageEnglishMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            LanguageFrenchMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            LanguageGermanMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            LanguageDutchMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            LanguageSpanishMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            LanguageSimpChineseMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            LanguageTradChineseMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            generateHyperlinksToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            groupBox1 = new System.Windows.Forms.GroupBox();
            webBrowser1 = new System.Windows.Forms.WebBrowser();
            toolStrip1 = new System.Windows.Forms.ToolStrip();
            PrintButton = new System.Windows.Forms.ToolStripButton();
            ForwardButton = new System.Windows.Forms.ToolStripButton();
            BackButton = new System.Windows.Forms.ToolStripButton();
            SaveAsTextButton = new System.Windows.Forms.ToolStripButton();
            statusStrip1.SuspendLayout();
            panel2.SuspendLayout();
            menuStrip1.SuspendLayout();
            groupBox1.SuspendLayout();
            toolStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // statusStrip1
            // 
            statusStrip1.ImageScalingSize = new System.Drawing.Size(32, 32);
            statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { StatusLabel });
            statusStrip1.Location = new System.Drawing.Point(0, 818);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Padding = new System.Windows.Forms.Padding(0, 0, 7, 0);
            statusStrip1.Size = new System.Drawing.Size(1148, 22);
            statusStrip1.TabIndex = 11;
            statusStrip1.Text = "Select a MSG file to open";
            // 
            // StatusLabel
            // 
            StatusLabel.Name = "StatusLabel";
            StatusLabel.Size = new System.Drawing.Size(178, 17);
            StatusLabel.Text = "Select a MSG or EML file to open";
            // 
            // panel2
            // 
            panel2.Controls.Add(FilesListBox);
            panel2.Controls.Add(label1);
            panel2.Controls.Add(menuStrip1);
            panel2.Dock = System.Windows.Forms.DockStyle.Top;
            panel2.Location = new System.Drawing.Point(0, 0);
            panel2.Margin = new System.Windows.Forms.Padding(2);
            panel2.Name = "panel2";
            panel2.Size = new System.Drawing.Size(1148, 94);
            panel2.TabIndex = 13;
            // 
            // FilesListBox
            // 
            FilesListBox.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            FilesListBox.FormattingEnabled = true;
            FilesListBox.ItemHeight = 15;
            FilesListBox.Location = new System.Drawing.Point(6, 22);
            FilesListBox.Margin = new System.Windows.Forms.Padding(2);
            FilesListBox.Name = "FilesListBox";
            FilesListBox.Size = new System.Drawing.Size(1138, 49);
            FilesListBox.TabIndex = 17;
            FilesListBox.DoubleClick += FilesListBox_DoubleClick;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(-79, 12);
            label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(83, 15);
            label1.TabIndex = 14;
            label1.Text = "Extracted files:";
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { fileToolStripMenuItem, settingsToolStripMenuItem });
            menuStrip1.Location = new System.Drawing.Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Padding = new System.Windows.Forms.Padding(3, 1, 0, 1);
            menuStrip1.Size = new System.Drawing.Size(1148, 24);
            menuStrip1.TabIndex = 22;
            menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { openToolStripMenuItem, toolStripSeparator1, printToolStripMenuItem, toolStripSeparator2, exitToolStripMenuItem });
            fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            fileToolStripMenuItem.Size = new System.Drawing.Size(37, 22);
            fileToolStripMenuItem.Text = "&File";
            // 
            // openToolStripMenuItem
            // 
            openToolStripMenuItem.Image = (System.Drawing.Image)resources.GetObject("openToolStripMenuItem.Image");
            openToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            openToolStripMenuItem.Name = "openToolStripMenuItem";
            openToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O;
            openToolStripMenuItem.Size = new System.Drawing.Size(146, 22);
            openToolStripMenuItem.Text = "&Open";
            openToolStripMenuItem.Click += openToolStripMenuItem_Click;
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            toolStripSeparator1.Size = new System.Drawing.Size(143, 6);
            // 
            // printToolStripMenuItem
            // 
            printToolStripMenuItem.Image = (System.Drawing.Image)resources.GetObject("printToolStripMenuItem.Image");
            printToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            printToolStripMenuItem.Name = "printToolStripMenuItem";
            printToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.P;
            printToolStripMenuItem.Size = new System.Drawing.Size(146, 22);
            printToolStripMenuItem.Text = "&Print";
            printToolStripMenuItem.Click += printToolStripMenuItem_Click;
            // 
            // toolStripSeparator2
            // 
            toolStripSeparator2.Name = "toolStripSeparator2";
            toolStripSeparator2.Size = new System.Drawing.Size(143, 6);
            // 
            // exitToolStripMenuItem
            // 
            exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            exitToolStripMenuItem.Size = new System.Drawing.Size(146, 22);
            exitToolStripMenuItem.Text = "E&xit";
            exitToolStripMenuItem.Click += exitToolStripMenuItem_Click;
            // 
            // settingsToolStripMenuItem
            // 
            settingsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { languageToolStripMenuItem, generateHyperlinksToolStripMenuItem });
            settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            settingsToolStripMenuItem.Size = new System.Drawing.Size(61, 22);
            settingsToolStripMenuItem.Text = "Settings";
            // 
            // languageToolStripMenuItem
            // 
            languageToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { LanguageEnglishMenuItem, LanguageFrenchMenuItem, LanguageGermanMenuItem, LanguageDutchMenuItem, LanguageSpanishMenuItem, LanguageSimpChineseMenuItem, LanguageTradChineseMenuItem });
            languageToolStripMenuItem.Name = "languageToolStripMenuItem";
            languageToolStripMenuItem.Size = new System.Drawing.Size(178, 22);
            languageToolStripMenuItem.Text = "&Language";
            // 
            // LanguageEnglishMenuItem
            // 
            LanguageEnglishMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            LanguageEnglishMenuItem.Name = "LanguageEnglishMenuItem";
            LanguageEnglishMenuItem.Size = new System.Drawing.Size(177, 22);
            LanguageEnglishMenuItem.Text = "English US (default)";
            LanguageEnglishMenuItem.Click += LanguageToolStripMenuItem_Click;
            // 
            // LanguageFrenchMenuItem
            // 
            LanguageFrenchMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            LanguageFrenchMenuItem.Name = "LanguageFrenchMenuItem";
            LanguageFrenchMenuItem.Size = new System.Drawing.Size(177, 22);
            LanguageFrenchMenuItem.Text = "French";
            LanguageFrenchMenuItem.Click += LanguageToolStripMenuItem_Click;
            // 
            // LanguageGermanMenuItem
            // 
            LanguageGermanMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            LanguageGermanMenuItem.Name = "LanguageGermanMenuItem";
            LanguageGermanMenuItem.Size = new System.Drawing.Size(177, 22);
            LanguageGermanMenuItem.Text = "German";
            LanguageGermanMenuItem.Click += LanguageToolStripMenuItem_Click;
            // 
            // LanguageDutchMenuItem
            // 
            LanguageDutchMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            LanguageDutchMenuItem.Name = "LanguageDutchMenuItem";
            LanguageDutchMenuItem.Size = new System.Drawing.Size(177, 22);
            LanguageDutchMenuItem.Text = "Dutch";
            LanguageDutchMenuItem.Click += LanguageToolStripMenuItem_Click;
            // 
            // LanguageSpanishMenuItem
            // 
            LanguageSpanishMenuItem.Name = "LanguageSpanishMenuItem";
            LanguageSpanishMenuItem.Size = new System.Drawing.Size(177, 22);
            LanguageSpanishMenuItem.Text = "Spanish";
            LanguageSpanishMenuItem.Click += LanguageToolStripMenuItem_Click;
            // 
            // LanguageSimpChineseMenuItem
            // 
            LanguageSimpChineseMenuItem.Name = "LanguageSimpChineseMenuItem";
            LanguageSimpChineseMenuItem.Size = new System.Drawing.Size(177, 22);
            LanguageSimpChineseMenuItem.Text = "Simp. Chinese";
            LanguageSimpChineseMenuItem.Click += LanguageToolStripMenuItem_Click;
            // 
            // LanguageTradChineseMenuItem
            // 
            LanguageTradChineseMenuItem.Name = "LanguageTradChineseMenuItem";
            LanguageTradChineseMenuItem.Size = new System.Drawing.Size(177, 22);
            LanguageTradChineseMenuItem.Text = "Trad. Chinese";
            LanguageTradChineseMenuItem.Click += LanguageToolStripMenuItem_Click;
            // 
            // generateHyperlinksToolStripMenuItem
            // 
            generateHyperlinksToolStripMenuItem.CheckOnClick = true;
            generateHyperlinksToolStripMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            generateHyperlinksToolStripMenuItem.Name = "generateHyperlinksToolStripMenuItem";
            generateHyperlinksToolStripMenuItem.Size = new System.Drawing.Size(178, 22);
            generateHyperlinksToolStripMenuItem.Text = "&Generate hyperlinks";
            generateHyperlinksToolStripMenuItem.Click += GenerateHyperlinksToolStripMenuItem_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(webBrowser1);
            groupBox1.Controls.Add(toolStrip1);
            groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            groupBox1.Location = new System.Drawing.Point(0, 94);
            groupBox1.Margin = new System.Windows.Forms.Padding(2);
            groupBox1.Name = "groupBox1";
            groupBox1.Padding = new System.Windows.Forms.Padding(2);
            groupBox1.Size = new System.Drawing.Size(1148, 724);
            groupBox1.TabIndex = 20;
            groupBox1.TabStop = false;
            groupBox1.Text = "Preview";
            // 
            // webBrowser1
            // 
            webBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
            webBrowser1.Location = new System.Drawing.Point(2, 43);
            webBrowser1.Margin = new System.Windows.Forms.Padding(2);
            webBrowser1.MinimumSize = new System.Drawing.Size(10, 10);
            webBrowser1.Name = "webBrowser1";
            webBrowser1.Size = new System.Drawing.Size(1144, 679);
            webBrowser1.TabIndex = 12;
            webBrowser1.Navigated += webBrowser1_Navigated_1;
            // 
            // toolStrip1
            // 
            toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { PrintButton, ForwardButton, BackButton, SaveAsTextButton });
            toolStrip1.Location = new System.Drawing.Point(2, 18);
            toolStrip1.Name = "toolStrip1";
            toolStrip1.Padding = new System.Windows.Forms.Padding(0);
            toolStrip1.Size = new System.Drawing.Size(1144, 25);
            toolStrip1.TabIndex = 0;
            toolStrip1.Text = "toolStrip1";
            // 
            // PrintButton
            // 
            PrintButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            PrintButton.Image = (System.Drawing.Image)resources.GetObject("PrintButton.Image");
            PrintButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            PrintButton.Name = "PrintButton";
            PrintButton.Size = new System.Drawing.Size(23, 22);
            PrintButton.Text = "&Print";
            PrintButton.Click += PrintButton_Click;
            // 
            // ForwardButton
            // 
            ForwardButton.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            ForwardButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            ForwardButton.Enabled = false;
            ForwardButton.Image = Properties.Resources.forward_icon;
            ForwardButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            ForwardButton.Name = "ForwardButton";
            ForwardButton.Size = new System.Drawing.Size(23, 22);
            ForwardButton.Text = "Go &forward";
            ForwardButton.Click += ForwardButton_Click_1;
            // 
            // BackButton
            // 
            BackButton.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            BackButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            BackButton.Enabled = false;
            BackButton.Image = Properties.Resources.back_icon;
            BackButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            BackButton.Name = "BackButton";
            BackButton.Size = new System.Drawing.Size(23, 22);
            BackButton.Text = "Go &back";
            BackButton.Click += BackButton_Click_1;
            // 
            // SaveAsTextButton
            // 
            SaveAsTextButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            SaveAsTextButton.Image = (System.Drawing.Image)resources.GetObject("SaveAsTextButton.Image");
            SaveAsTextButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            SaveAsTextButton.Name = "SaveAsTextButton";
            SaveAsTextButton.Size = new System.Drawing.Size(23, 22);
            SaveAsTextButton.Text = "Save as text";
            SaveAsTextButton.Click += SaveAsTextButton_Click;
            // 
            // ViewerForm
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            ClientSize = new System.Drawing.Size(1148, 840);
            Controls.Add(groupBox1);
            Controls.Add(panel2);
            Controls.Add(statusStrip1);
            settings1.ClientSize = new System.Drawing.Size(1062, 903);
            settings1.GenerateHyperLinks = true;
            settings1.InitialDirectory = "";
            settings1.Language = 1;
            settings1.Location = new System.Drawing.Point(0, 0);
            settings1.Placement = "";
            settings1.SaveDirectory = "";
            settings1.SettingsKey = "";
            settings1.ShowMessageHeader = true;
            settings1.WindowState = System.Windows.Forms.FormWindowState.Normal;
            DataBindings.Add(new System.Windows.Forms.Binding("WindowState", settings1, "WindowState", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
            Margin = new System.Windows.Forms.Padding(2);
            Name = "ViewerForm";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Load += ViewerForm_Load;
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            panel2.ResumeLayout(false);
            panel2.PerformLayout();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            toolStrip1.ResumeLayout(false);
            toolStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }
        #endregion

        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox FilesListBox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton BackButton;
        private System.Windows.Forms.ToolStripButton ForwardButton;
        private System.Windows.Forms.ToolStripButton PrintButton;
        private System.Windows.Forms.ToolStripStatusLabel StatusLabel;
        private System.Windows.Forms.ToolStripButton SaveAsTextButton;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem printToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem settingsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem languageToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem LanguageEnglishMenuItem;
        private System.Windows.Forms.ToolStripMenuItem LanguageFrenchMenuItem;
        private System.Windows.Forms.ToolStripMenuItem LanguageGermanMenuItem;
        private System.Windows.Forms.ToolStripMenuItem LanguageDutchMenuItem;
        private System.Windows.Forms.ToolStripMenuItem generateHyperlinksToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem LanguageSimpChineseMenuItem;
        private System.Windows.Forms.ToolStripMenuItem LanguageSpanishMenuItem;
        private System.Windows.Forms.ToolStripMenuItem LanguageTradChineseMenuItem;
    }
}

