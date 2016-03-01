using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Timers;
using System.Reflection;
using System.IO;

namespace GmailNotifier
{
    public class Program : Form
    {
        [STAThread]
        public static void Main()
        {
            Application.Run(new Program());
        }
 
        private NotifyIcon trayIcon;
        private ContextMenu trayMenu;
		private MenuItem menuItem;

		private System.Timers.Timer timer;
		private GmailService gmail;
		private int lastCount = 0;
		private const int timeoutStep = 60000; // minute increments
		private const int timeoutMax = 300000; // 5 minutes wait max

        public Program()
        {
			// create a timer
			timer = new System.Timers.Timer(1000);
			timer.Elapsed += OnTimedEvent;

			menuItem = new MenuItem("Launch on Windows Start", OnToggleAutoStart) { Checked = AppIsAutostart() };

            // create tray menu
            trayMenu = new ContextMenu();
			trayMenu.MenuItems.Add("Open Gmail", OnOpenGmail);
			trayMenu.MenuItems.Add("Re-authenticate", OnReAuth);
			trayMenu.MenuItems.Add(menuItem);
            trayMenu.MenuItems.Add("Exit", OnExit);

            // create a tray icon
            trayIcon = new NotifyIcon();
            trayIcon.Text = "Gmail Notifier: Starting";
			trayIcon.Icon = Icon.FromHandle(Properties.Resources.black_icon.GetHicon()); // new Icon(SystemIcons.Application, 40, 40);
            trayIcon.ContextMenu = trayMenu;
			trayIcon.DoubleClick += OnOpenGmail;
            trayIcon.Visible = true;
        }
 
        protected override void OnLoad(EventArgs e)
        {
            Visible = false;		// Hide form window
            ShowInTaskbar = false;	// Remove from taskbar
            base.OnLoad(e);

			// create gmail service and start timer
			CreateGmailService();
        }

        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
				try { timer.Dispose(); } catch { }
				try { gmail.Dispose(); } catch { }
				trayIcon.Dispose();
            }
 
            base.Dispose(isDisposing);
        }
		 
        private void OnExit(object sender, EventArgs e)
        {
            Application.Exit();
        }
 
        private void OnReAuth(object sender, EventArgs e)
        {
			timer.Stop();

			try { gmail.Dispose(); } catch { }

            string credPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal);
			credPath = Path.Combine(credPath, ".credentials/ply-gmail-notifier.json");
			try { Directory.Delete(credPath, true); } catch { }

			CreateGmailService();
        }

		private void OnToggleAutoStart(object sender, EventArgs e)
		{
			Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
			if (k.GetValue("plyGmailNotifier") == null)
			{
				menuItem.Checked = true;
				k.SetValue("plyGmailNotifier", Application.ExecutablePath.ToString());
			}
			else
			{
				menuItem.Checked = false;
				k.DeleteValue("plyGmailNotifier");
			}
		}

		private bool AppIsAutostart()
		{
			Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
			return (k.GetValue("plyGmailNotifier") != null);
		}

		private void OnOpenGmail(object sender, EventArgs e)
		{
			System.Diagnostics.Process.Start("http://mail.google.com/");
		}

		private void OnTimedEvent(Object source, ElapsedEventArgs e)
		{
			// stop timer wile checking mail
			timer.Stop();

			// check mail
			try
			{ 
				CheckInbox();
			} 
			catch 
			{ 
				// this can happen if a firewall blocks the app while waiting for user to allow
				// for example: System.Net.Http.HttpRequestException
				timer.Interval = 60000; // try again in a minute
				timer.Start();
				return;
			}

			// restart timer, it will wait longer and longer until max timeout is reached
			// if no new mails arrived the timeout is reset by CheckInbox() 
			if (timer.Interval < timeoutMax) timer.Interval += timeoutStep;
			timer.Start();
		}

		private void CreateGmailService()
		{
            trayIcon.Text = "Gmail Notifier: Authenticating";
			trayIcon.Icon = Icon.FromHandle(Properties.Resources.black_icon.GetHicon());

			string[] Scopes = { GmailService.Scope.GmailReadonly };
			string ApplicationName = "Gmail Notifier";

			// get credentials
			UserCredential credential;
			using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("GmailNotifier.client_id.json"))
            {
                string credPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/ply-gmail-notifier.json");
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
            }

			// create Gmail API service
			gmail = new GmailService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName
			});

			// start timer
			trayIcon.Text = "0 unread message";
			timer.Interval = 1000; // wait a minute before checking mail
			timer.Start();
		}

		private void CheckInbox()
		{
			Google.Apis.Gmail.v1.Data.Label label = gmail.Users.Labels.Get("me", "INBOX").Execute();
			int count = label.MessagesUnread ?? default(int);

			if (count != lastCount)
			{
				trayIcon.Text = count.ToString() + " unread message";
				if (count > 0) trayIcon.Icon = Icon.FromHandle(Properties.Resources.red_icon.GetHicon());
				else trayIcon.Icon = Icon.FromHandle(Properties.Resources.black_icon.GetHicon());

				if (count > lastCount)
				{
					DoNotification("New message." + Environment.NewLine + count.ToString() + " unread message.");
				}

				lastCount = count;

				// reset timer when mail count changed
				timer.Interval = 1;
			}
		}

		private void DoNotification(string message)
		{
			trayIcon.Visible = true;
			trayIcon.BalloonTipIcon = ToolTipIcon.None;
			trayIcon.BalloonTipTitle = "Gmail Notifier";
			trayIcon.BalloonTipText = message;
			trayIcon.ShowBalloonTip(10000); // note: timeout is actually ignored
		}

    }
}
