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
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.StatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.HyperLinkCheckBox = new System.Windows.Forms.CheckBox();
            this.SelectButton = new System.Windows.Forms.Button();
            this.FilesListBox = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.PrintButton = new System.Windows.Forms.ToolStripButton();
            this.ForwardButton = new System.Windows.Forms.ToolStripButton();
            this.BackButton = new System.Windows.Forms.ToolStripButton();
            this.SaveAsTextButton = new System.Windows.Forms.ToolStripButton();
            this.statusStrip1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.StatusLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 866);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1062, 37);
            this.statusStrip1.TabIndex = 11;
            this.statusStrip1.Text = "Select a MSG file to open";
            // 
            // StatusLabel
            // 
            this.StatusLabel.Name = "StatusLabel";
            this.StatusLabel.Size = new System.Drawing.Size(287, 32);
            this.StatusLabel.Text = "Select a MSG file to open";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.HyperLinkCheckBox);
            this.panel2.Controls.Add(this.SelectButton);
            this.panel2.Controls.Add(this.FilesListBox);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1062, 236);
            this.panel2.TabIndex = 13;
            // 
            // HyperLinkCheckBox
            // 
            this.HyperLinkCheckBox.AutoSize = true;
            this.HyperLinkCheckBox.Checked = global::MsgViewer.Properties.Settings.Default.GenereateHyperLinks;
            this.HyperLinkCheckBox.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::MsgViewer.Properties.Settings.Default, "GenereateHyperLinks", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.HyperLinkCheckBox.Location = new System.Drawing.Point(288, 20);
            this.HyperLinkCheckBox.Name = "HyperLinkCheckBox";
            this.HyperLinkCheckBox.Size = new System.Drawing.Size(237, 29);
            this.HyperLinkCheckBox.TabIndex = 21;
            this.HyperLinkCheckBox.Text = "Generate hyperlinks";
            this.HyperLinkCheckBox.UseVisualStyleBackColor = true;
            // 
            // SelectButton
            // 
            this.SelectButton.Location = new System.Drawing.Point(12, 12);
            this.SelectButton.Name = "SelectButton";
            this.SelectButton.Size = new System.Drawing.Size(257, 42);
            this.SelectButton.TabIndex = 20;
            this.SelectButton.Text = "Select MSG or EML file";
            this.SelectButton.UseVisualStyleBackColor = true;
            this.SelectButton.Click += new System.EventHandler(this.SelectButton_Click);
            // 
            // FilesListBox
            // 
            this.FilesListBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.FilesListBox.FormattingEnabled = true;
            this.FilesListBox.ItemHeight = 25;
            this.FilesListBox.Location = new System.Drawing.Point(12, 70);
            this.FilesListBox.Name = "FilesListBox";
            this.FilesListBox.Size = new System.Drawing.Size(1038, 154);
            this.FilesListBox.TabIndex = 17;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(-158, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(154, 25);
            this.label1.TabIndex = 14;
            this.label1.Text = "Extracted files:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.webBrowser1);
            this.groupBox1.Controls.Add(this.toolStrip1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 236);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(1062, 630);
            this.groupBox1.TabIndex = 20;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Preview";
            // 
            // webBrowser1
            // 
            this.webBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser1.Location = new System.Drawing.Point(3, 66);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(1056, 561);
            this.webBrowser1.TabIndex = 12;
            this.webBrowser1.Navigated += new System.Windows.Forms.WebBrowserNavigatedEventHandler(this.webBrowser1_Navigated_1);
            // 
            // toolStrip1
            // 
            this.toolStrip1.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.PrintButton,
            this.ForwardButton,
            this.BackButton,
            this.SaveAsTextButton});
            this.toolStrip1.Location = new System.Drawing.Point(3, 27);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1056, 39);
            this.toolStrip1.TabIndex = 0;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // PrintButton
            // 
            this.PrintButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.PrintButton.Image = ((System.Drawing.Image)(resources.GetObject("PrintButton.Image")));
            this.PrintButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.PrintButton.Name = "PrintButton";
            this.PrintButton.Size = new System.Drawing.Size(36, 36);
            this.PrintButton.Text = "&Print";
            this.PrintButton.Click += new System.EventHandler(this.PrintButton_Click);
            // 
            // ForwardButton
            // 
            this.ForwardButton.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.ForwardButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ForwardButton.Image = global::MsgViewer.Properties.Resources.forward_icon;
            this.ForwardButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ForwardButton.Name = "ForwardButton";
            this.ForwardButton.Size = new System.Drawing.Size(36, 36);
            this.ForwardButton.Text = "Go &forward";
            this.ForwardButton.Click += new System.EventHandler(this.ForwardButton_Click_1);
            // 
            // BackButton
            // 
            this.BackButton.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.BackButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.BackButton.Image = global::MsgViewer.Properties.Resources.back_icon;
            this.BackButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.BackButton.Name = "BackButton";
            this.BackButton.Size = new System.Drawing.Size(36, 36);
            this.BackButton.Text = "Go &back";
            this.BackButton.Click += new System.EventHandler(this.BackButton_Click_1);
            // 
            // SaveAsTextButton
            // 
            this.SaveAsTextButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.SaveAsTextButton.Image = ((System.Drawing.Image)(resources.GetObject("SaveAsTextButton.Image")));
            this.SaveAsTextButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.SaveAsTextButton.Name = "SaveAsTextButton";
            this.SaveAsTextButton.Size = new System.Drawing.Size(36, 36);
            this.SaveAsTextButton.Text = "Save as text";
            this.SaveAsTextButton.Click += new System.EventHandler(this.SaveAsTextButton_Click);
            // 
            // ViewerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(192F, 192F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(1062, 903);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.statusStrip1);
            this.DataBindings.Add(new System.Windows.Forms.Binding("WindowState", global::MsgViewer.Properties.Settings.Default, "WindowState", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ViewerForm";
            this.Text = "MSG Viewer v1.9.3";
            this.WindowState = global::MsgViewer.Properties.Settings.Default.WindowState;
            this.Load += new System.EventHandler(this.ViewerForm_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
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
        private System.Windows.Forms.Button SelectButton;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton BackButton;
        private System.Windows.Forms.ToolStripButton ForwardButton;
        private System.Windows.Forms.ToolStripButton PrintButton;
        private System.Windows.Forms.ToolStripStatusLabel StatusLabel;
        private System.Windows.Forms.CheckBox HyperLinkCheckBox;
        private System.Windows.Forms.ToolStripButton SaveAsTextButton;

    }
}

