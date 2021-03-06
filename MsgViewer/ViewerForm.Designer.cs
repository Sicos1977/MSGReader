﻿namespace MsgViewer
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
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.StatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.FilesListBox = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.printToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.settingsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.languageToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.LanguageEnglishMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.LanguageFrenchMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.LanguageGermanMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.LanguageDutchMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.LanguageSimpChineseMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.LanguageSpanishMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.LanguageTradChineseMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.generateHyperlinksToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.PrintButton = new System.Windows.Forms.ToolStripButton();
            this.ForwardButton = new System.Windows.Forms.ToolStripButton();
            this.BackButton = new System.Windows.Forms.ToolStripButton();
            this.SaveAsTextButton = new System.Windows.Forms.ToolStripButton();
            this.statusStrip1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.StatusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 430);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(0, 0, 7, 0);
            this.statusStrip1.Size = new System.Drawing.Size(531, 22);
            this.statusStrip1.TabIndex = 11;
            this.statusStrip1.Text = "Select a MSG file to open";
            // 
            // StatusLabel
            // 
            this.StatusLabel.Name = "StatusLabel";
            this.StatusLabel.Size = new System.Drawing.Size(178, 17);
            this.StatusLabel.Text = "Select a MSG or EML file to open";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.FilesListBox);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.menuStrip1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Margin = new System.Windows.Forms.Padding(2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(531, 94);
            this.panel2.TabIndex = 13;
            // 
            // FilesListBox
            // 
            this.FilesListBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.FilesListBox.FormattingEnabled = true;
            this.FilesListBox.Location = new System.Drawing.Point(6, 22);
            this.FilesListBox.Margin = new System.Windows.Forms.Padding(2);
            this.FilesListBox.Name = "FilesListBox";
            this.FilesListBox.Size = new System.Drawing.Size(521, 56);
            this.FilesListBox.TabIndex = 17;
            this.FilesListBox.DoubleClick += new System.EventHandler(this.FilesListBox_DoubleClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(-79, 12);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 13);
            this.label1.TabIndex = 14;
            this.label1.Text = "Extracted files:";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.settingsToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Padding = new System.Windows.Forms.Padding(3, 1, 0, 1);
            this.menuStrip1.Size = new System.Drawing.Size(531, 24);
            this.menuStrip1.TabIndex = 22;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.toolStripSeparator1,
            this.printToolStripMenuItem,
            this.toolStripSeparator2,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 22);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("openToolStripMenuItem.Image")));
            this.openToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.openToolStripMenuItem.Size = new System.Drawing.Size(146, 22);
            this.openToolStripMenuItem.Text = "&Open";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(143, 6);
            // 
            // printToolStripMenuItem
            // 
            this.printToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("printToolStripMenuItem.Image")));
            this.printToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.printToolStripMenuItem.Name = "printToolStripMenuItem";
            this.printToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.P)));
            this.printToolStripMenuItem.Size = new System.Drawing.Size(146, 22);
            this.printToolStripMenuItem.Text = "&Print";
            this.printToolStripMenuItem.Click += new System.EventHandler(this.printToolStripMenuItem_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(143, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(146, 22);
            this.exitToolStripMenuItem.Text = "E&xit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // settingsToolStripMenuItem
            // 
            this.settingsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.languageToolStripMenuItem,
            this.generateHyperlinksToolStripMenuItem});
            this.settingsToolStripMenuItem.Name = "settingsToolStripMenuItem";
            this.settingsToolStripMenuItem.Size = new System.Drawing.Size(61, 22);
            this.settingsToolStripMenuItem.Text = "Settings";
            // 
            // languageToolStripMenuItem
            // 
            this.languageToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.LanguageEnglishMenuItem,
            this.LanguageFrenchMenuItem,
            this.LanguageGermanMenuItem,
            this.LanguageDutchMenuItem,
            this.LanguageSpanishMenuItem,
            this.LanguageSimpChineseMenuItem,
            this.LanguageTradChineseMenuItem});
            this.languageToolStripMenuItem.Name = "languageToolStripMenuItem";
            this.languageToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.languageToolStripMenuItem.Text = "&Language";
            // 
            // LanguageEnglishMenuItem
            // 
            this.LanguageEnglishMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.LanguageEnglishMenuItem.Name = "LanguageEnglishMenuItem";
            this.LanguageEnglishMenuItem.Size = new System.Drawing.Size(180, 22);
            this.LanguageEnglishMenuItem.Text = "English US (default)";
            this.LanguageEnglishMenuItem.Click += new System.EventHandler(this.LanguageToolStripMenuItem_Click);
            // 
            // LanguageFrenchMenuItem
            // 
            this.LanguageFrenchMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.LanguageFrenchMenuItem.Name = "LanguageFrenchMenuItem";
            this.LanguageFrenchMenuItem.Size = new System.Drawing.Size(180, 22);
            this.LanguageFrenchMenuItem.Text = "French";
            this.LanguageFrenchMenuItem.Click += new System.EventHandler(this.LanguageToolStripMenuItem_Click);
            // 
            // LanguageGermanMenuItem
            // 
            this.LanguageGermanMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.LanguageGermanMenuItem.Name = "LanguageGermanMenuItem";
            this.LanguageGermanMenuItem.Size = new System.Drawing.Size(180, 22);
            this.LanguageGermanMenuItem.Text = "German";
            this.LanguageGermanMenuItem.Click += new System.EventHandler(this.LanguageToolStripMenuItem_Click);
            // 
            // LanguageDutchMenuItem
            // 
            this.LanguageDutchMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.LanguageDutchMenuItem.Name = "LanguageDutchMenuItem";
            this.LanguageDutchMenuItem.Size = new System.Drawing.Size(180, 22);
            this.LanguageDutchMenuItem.Text = "Dutch";
            this.LanguageDutchMenuItem.Click += new System.EventHandler(this.LanguageToolStripMenuItem_Click);
            // 
            // LanguageSimpChineseMenuItem
            // 
            this.LanguageSimpChineseMenuItem.Name = "LanguageSimpChineseMenuItem";
            this.LanguageSimpChineseMenuItem.Size = new System.Drawing.Size(180, 22);
            this.LanguageSimpChineseMenuItem.Text = "Simp. Chinese";
            this.LanguageSimpChineseMenuItem.Click += new System.EventHandler(this.LanguageToolStripMenuItem_Click);
            // 
            // LanguageSpanishMenuItem
            // 
            this.LanguageSpanishMenuItem.Name = "LanguageSpanishMenuItem";
            this.LanguageSpanishMenuItem.Size = new System.Drawing.Size(180, 22);
            this.LanguageSpanishMenuItem.Text = "Spanish";
            this.LanguageSpanishMenuItem.Click += new System.EventHandler(this.LanguageToolStripMenuItem_Click);
            // 
            // LanguageTradChineseMenuItem
            // 
            this.LanguageTradChineseMenuItem.Name = "LanguageTradChineseMenuItem";
            this.LanguageTradChineseMenuItem.Size = new System.Drawing.Size(180, 22);
            this.LanguageTradChineseMenuItem.Text = "Trad. Chinese";
            this.LanguageTradChineseMenuItem.Click += new System.EventHandler(this.LanguageToolStripMenuItem_Click);
            // 
            // generateHyperlinksToolStripMenuItem
            // 
            this.generateHyperlinksToolStripMenuItem.CheckOnClick = true;
            this.generateHyperlinksToolStripMenuItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.generateHyperlinksToolStripMenuItem.Name = "generateHyperlinksToolStripMenuItem";
            this.generateHyperlinksToolStripMenuItem.Size = new System.Drawing.Size(184, 22);
            this.generateHyperlinksToolStripMenuItem.Text = "&Generate hyperlinks";
            this.generateHyperlinksToolStripMenuItem.Click += new System.EventHandler(this.GenerateHyperlinksToolStripMenuItem_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.webBrowser1);
            this.groupBox1.Controls.Add(this.toolStrip1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 94);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(531, 336);
            this.groupBox1.TabIndex = 20;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Preview";
            // 
            // webBrowser1
            // 
            this.webBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser1.Location = new System.Drawing.Point(2, 40);
            this.webBrowser1.Margin = new System.Windows.Forms.Padding(2);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(10, 10);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(527, 294);
            this.webBrowser1.TabIndex = 12;
            this.webBrowser1.Navigated += new System.Windows.Forms.WebBrowserNavigatedEventHandler(this.webBrowser1_Navigated_1);
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.PrintButton,
            this.ForwardButton,
            this.BackButton,
            this.SaveAsTextButton});
            this.toolStrip1.Location = new System.Drawing.Point(2, 15);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Padding = new System.Windows.Forms.Padding(0);
            this.toolStrip1.Size = new System.Drawing.Size(527, 25);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // PrintButton
            // 
            this.PrintButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.PrintButton.Image = ((System.Drawing.Image)(resources.GetObject("PrintButton.Image")));
            this.PrintButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.PrintButton.Name = "PrintButton";
            this.PrintButton.Size = new System.Drawing.Size(23, 22);
            this.PrintButton.Text = "&Print";
            this.PrintButton.Click += new System.EventHandler(this.PrintButton_Click);
            // 
            // ForwardButton
            // 
            this.ForwardButton.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.ForwardButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ForwardButton.Enabled = false;
            this.ForwardButton.Image = global::MsgViewer.Properties.Resources.forward_icon;
            this.ForwardButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ForwardButton.Name = "ForwardButton";
            this.ForwardButton.Size = new System.Drawing.Size(23, 22);
            this.ForwardButton.Text = "Go &forward";
            this.ForwardButton.Click += new System.EventHandler(this.ForwardButton_Click_1);
            // 
            // BackButton
            // 
            this.BackButton.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.BackButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.BackButton.Enabled = false;
            this.BackButton.Image = global::MsgViewer.Properties.Resources.back_icon;
            this.BackButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.BackButton.Name = "BackButton";
            this.BackButton.Size = new System.Drawing.Size(23, 22);
            this.BackButton.Text = "Go &back";
            this.BackButton.Click += new System.EventHandler(this.BackButton_Click_1);
            // 
            // SaveAsTextButton
            // 
            this.SaveAsTextButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.SaveAsTextButton.Image = ((System.Drawing.Image)(resources.GetObject("SaveAsTextButton.Image")));
            this.SaveAsTextButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.SaveAsTextButton.Name = "SaveAsTextButton";
            this.SaveAsTextButton.Size = new System.Drawing.Size(23, 22);
            this.SaveAsTextButton.Text = "Save as text";
            this.SaveAsTextButton.Click += new System.EventHandler(this.SaveAsTextButton_Click);
            // 
            // ViewerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(531, 452);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.statusStrip1);
            this.DataBindings.Add(new System.Windows.Forms.Binding("WindowState", global::MsgViewer.Properties.Settings.Default, "WindowState", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "ViewerForm";
            this.WindowState = global::MsgViewer.Properties.Settings.Default.WindowState;
            this.Load += new System.EventHandler(this.ViewerForm_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

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

