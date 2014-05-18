using System;
using System.Drawing;
using System.Windows.Forms;

namespace MsgMapper
{
    public partial class MainForm : Form
    {
        #region Fields
        private readonly NotifyIcon _trayIcon;
        #endregion

        public MainForm()
        {
            InitializeComponent();

            var trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Show", OnShow);
            trayMenu.MenuItems.Add("Exit", OnExit);

            // Create a tray icon
            _trayIcon = new NotifyIcon
            {
                Text = "MSG properties to extended file properties mapper",
                Icon = new Icon(SystemIcons.Application, 40, 40),
                ContextMenu = trayMenu,
                Visible = true
            };
        }
 
        protected override void OnLoad(EventArgs e)
        {
            Visible       = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.
 
            base.OnLoad(e);
        }
 
        private void OnExit(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void OnShow(object sender, EventArgs e)
        {
            Show();
        }
    }
}
