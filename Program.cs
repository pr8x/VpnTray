using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using DotRas;

namespace VpnTray
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Application.Run(new MyCustomApplicationContext());
        }
    }

    public class MyCustomApplicationContext : ApplicationContext
    {
        private const int RAS_MaxEntryName = 256;
        private const int MAX_PATH = 260;
        private readonly Image _checkIcon;

        private readonly Icon _offIcon;
        private readonly Icon _onIcon;
        private readonly NotifyIcon _trayIcon;
        private readonly RasConnectionWatcher _watcher;


        public MyCustomApplicationContext()
        {
            MainForm = new Form
            {
                ShowInTaskbar = false, 
                WindowState = FormWindowState.Minimized
            };

            _watcher = new RasConnectionWatcher();
            _watcher.Connected += OnConnection;
            _watcher.Disconnected += OnDisconnected;
            _watcher.Start();

            _checkIcon = Image.FromFile("check.png");
            _offIcon = IconFromImage(Image.FromFile("off.png"));
            _onIcon = IconFromImage(Image.FromFile("on.png"));

            _trayIcon = new NotifyIcon
            {
                Visible = true,
                ContextMenuStrip = new ContextMenuStrip()
            };

            UpdateTrayIcon();
        }

        private static Icon IconFromImage(Image img)
        {
            var ms = new MemoryStream();
            var bw = new BinaryWriter(ms);
            // Header
            bw.Write((short) 0); // 0 : reserved
            bw.Write((short) 1); // 2 : 1=ico, 2=cur
            bw.Write((short) 1); // 4 : number of images
            // Image directory
            var w = img.Width;
            if (w >= 256) w = 0;
            bw.Write((byte) w); // 0 : width of image
            var h = img.Height;
            if (h >= 256) h = 0;
            bw.Write((byte) h); // 1 : height of image
            bw.Write((byte) 0); // 2 : number of colors in palette
            bw.Write((byte) 0); // 3 : reserved
            bw.Write((short) 0); // 4 : number of color planes
            bw.Write((short) 0); // 6 : bits per pixel
            var sizeHere = ms.Position;
            bw.Write(0); // 8 : image size
            var start = (int) ms.Position + 4;
            bw.Write(start); // 12: offset of image data
            // Image data
            img.Save(ms, ImageFormat.Png);
            var imageSize = (int) ms.Position - start;
            ms.Seek(sizeHere, SeekOrigin.Begin);
            bw.Write(imageSize);
            ms.Seek(0, SeekOrigin.Begin);

            // And load it
            return new Icon(ms);
        }

        private void UpdateTrayIcon()
        {
            var connections = RasConnection.EnumerateConnections().ToList();

            if (connections.Count != 0)
            {
                _trayIcon.Text = $"Connected to: {string.Join(", ", connections.Select(x => x.EntryName))}";
                _trayIcon.Icon = _onIcon;
            }
            else
            {
                _trayIcon.Text = "Not connected";
                _trayIcon.Icon = _offIcon;
            }

            var strip = _trayIcon.ContextMenuStrip;

            strip.Items.Clear();

            foreach (var profile in GetProfiles())
            {
                var connection = connections.FirstOrDefault(x => x.EntryName.Equals(profile));

                strip.Items.Add(
                    profile,
                    connection != null ? _checkIcon : null,
                    (s, e) =>
                    {
                        if (connection != null)
                        {
                            connection.Disconnect(new CancellationToken());
                        }
                        else
                        {
                            var dialer = new RasDialer {EntryName = profile};

                            dialer.Connect();
                        }
                    });
            }

            strip.Items.Add(new ToolStripSeparator());

            strip.Items.Add("Exit", null, OnExit);
        }

        private void OnDisconnected(object? sender, RasConnectionEventArgs e)
        {
            MainForm!.Invoke((MethodInvoker) delegate
            {
                UpdateTrayIcon();
                _trayIcon.ShowBalloonTip(
                    2000,
                    null,
                    $"Disconnected from {e.ConnectionInformation.EntryName}",
                    ToolTipIcon.Info);
            });
        }

        private void OnConnection(object? sender, RasConnectionEventArgs e)
        {
            MainForm!.Invoke((MethodInvoker) delegate
            {
                UpdateTrayIcon();
                _trayIcon.ShowBalloonTip(
                    2000,
                    null,
                    $"Connected to {e.ConnectionInformation.EntryName}",
                    ToolTipIcon.Info);
            });
        }

        [DllImport("rasapi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern uint RasEnumEntries(IntPtr reserved, IntPtr lpszPhonebook,
            [In] [Out] RASENTRYNAME[] lprasentryname, ref int lpcb, ref int lpcEntries);


        private void OnExit(object? sender, EventArgs args)
        {
            _trayIcon.Visible = false;
            _watcher.Stop();

            Application.Exit();
        }

        public IEnumerable<string> GetProfiles()
        {
            int cb = Marshal.SizeOf(typeof(RASENTRYNAME)), entries = 0;
            var entryNames = new RASENTRYNAME[1];
            entryNames[0].dwSize = Marshal.SizeOf(typeof(RASENTRYNAME));

            var nRet = RasEnumEntries(IntPtr.Zero, IntPtr.Zero, entryNames, ref cb, ref entries);

            if (entries == 0)
            {
                yield break;
            }

            entryNames = new RASENTRYNAME[entries];

            for (var i = 0; i < entries; i++)
            {
                entryNames[i].dwSize = Marshal.SizeOf(typeof(RASENTRYNAME));
            }

            nRet = RasEnumEntries(IntPtr.Zero, IntPtr.Zero, entryNames, ref cb, ref entries);
            for (var i = 0; i < entries; i++)
            {
                yield return entryNames[i].szEntryName;
            }
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct RASENTRYNAME
        {
            public int dwSize;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = RAS_MaxEntryName + 1)]
            public readonly string szEntryName;

            public readonly int dwFlags;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = MAX_PATH + 1)]
            public readonly string szPhonebook;
        }
    }
}