namespace MsgMapper
{
    partial class MainForm
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
            if (disposing)
            {
                // Release the icon resource.
                _trayIcon.Dispose();
            }

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.AddFolderButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.StatusListBox = new System.Windows.Forms.ListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.FoldersDataGridView = new System.Windows.Forms.DataGridView();
            this.Folder = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.FoldersDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // AddFolderButton
            // 
            this.AddFolderButton.Location = new System.Drawing.Point(739, 213);
            this.AddFolderButton.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.AddFolderButton.Name = "AddFolderButton";
            this.AddFolderButton.Size = new System.Drawing.Size(70, 26);
            this.AddFolderButton.TabIndex = 0;
            this.AddFolderButton.Text = "Add folder";
            this.AddFolderButton.UseVisualStyleBackColor = true;
            this.AddFolderButton.Click += new System.EventHandler(this.AddFolderButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 5);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Folders to watch:";
            // 
            // StatusListBox
            // 
            this.StatusListBox.FormattingEnabled = true;
            this.StatusListBox.Location = new System.Drawing.Point(8, 260);
            this.StatusListBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.StatusListBox.Name = "StatusListBox";
            this.StatusListBox.Size = new System.Drawing.Size(801, 160);
            this.StatusListBox.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 246);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(40, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Status:";
            // 
            // label3
            // 
            this.label3.ForeColor = System.Drawing.Color.Red;
            this.label3.Location = new System.Drawing.Point(11, 220);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(661, 15);
            this.label3.TabIndex = 5;
            this.label3.Text = "This program uses a filesystemwatcher to monitor folders for Outlook Message File" +
    "s. Click the [Add Folder] button to add a folder to watch.";
            // 
            // FoldersDataGridView
            // 
            this.FoldersDataGridView.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.FoldersDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.FoldersDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Folder});
            this.FoldersDataGridView.Location = new System.Drawing.Point(8, 19);
            this.FoldersDataGridView.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.FoldersDataGridView.MultiSelect = false;
            this.FoldersDataGridView.Name = "FoldersDataGridView";
            this.FoldersDataGridView.ReadOnly = true;
            this.FoldersDataGridView.RowTemplate.Height = 33;
            this.FoldersDataGridView.Size = new System.Drawing.Size(801, 184);
            this.FoldersDataGridView.TabIndex = 1;
            // 
            // Folder
            // 
            this.Folder.HeaderText = "Folder";
            this.Folder.Name = "Folder";
            this.Folder.ReadOnly = true;
            this.Folder.Width = 740;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(820, 426);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.StatusListBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.FoldersDataGridView);
            this.Controls.Add(this.AddFolderButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MSG properties mapper";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.FoldersDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.Button AddFolderButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox StatusListBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView FoldersDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn Folder;
    }
}

