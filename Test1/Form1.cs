using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using NTwain;
using NTwain.Data;
using System.Drawing;
using System.Security.Principal;
using System.Windows.Shapes;

namespace Test1
{
    public partial class Form1 : Form
    {
        private Button btnInit;
        private ListBox lstStatus;
        private PictureBox picPreview;
        TwainSession _session;
        int _imageCount = 0;
        string _saveDir = @"C:\Scans";

        public Form1()
        {
            InitializeComponent();
            Button btn = new() { Text = "Scan", Location = new(20, 20), Size = new(100, 40) };
            btn.Click += Btn_Click;
            Controls.Add(btn);
            Size = new(400, 200);
        }

        private void InitializeUI()
        {
            Text = "Auto Duplex Scanner (.NET 8)";
            Size = new Size(800, 600);
            StartPosition = FormStartPosition.CenterScreen;

            btnInit = new Button
            {
                Text = "Initialize Scanner Watcher",
                Location = new Point(20, 20),
                Size = new Size(220, 40)
            };
            btnInit.Click += BtnInit_Click;
            Controls.Add(btnInit);

            lstStatus = new ListBox
            {
                Location = new Point(20, 80),
                Size = new Size(350, 450)
            };
            Controls.Add(lstStatus);

            picPreview = new PictureBox
            {
                Location = new Point(400, 80),
                Size = new Size(360, 450),
                BorderStyle = BorderStyle.FixedSingle,
                SizeMode = PictureBoxSizeMode.Zoom
            };
            Controls.Add(picPreview);
        }

        private void BtnInit_Click(object? sender, EventArgs e)
        {
            _ = Task.Run(InitScannerAsync);
        }

        private void Btn_Click(object sender, EventArgs e)
        {
            ThreadPool.QueueUserWorkItem(_ =>
            {
                if (!Directory.Exists(_saveDir)) Directory.CreateDirectory(_saveDir);

                var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, typeof(Form1).Assembly);
                _session = new TwainSession(appId);

                _session.Open();

                if (_session.State < 3)
                {
                    MessageBox.Show("Failed to open TWAIN session.");
                    return;
                }

                var sources = _session;
                var sourceList = _session.Select(s => s.Name).ToList();

                if (!sources.Any())
                {
                    MessageBox.Show("No scanners detected.");
                    return;
                }

                var source = sources.FirstOrDefault(s => s.Name.Contains("DS-530", StringComparison.OrdinalIgnoreCase)) ?? sources.First();
                source.Open();

                // Set duplex capabilities
                source.Capabilities.CapDuplexEnabled?.SetValue(BoolType.True);
                source.Capabilities.CapFeederEnabled?.SetValue(BoolType.True);
                source.Capabilities.CapAutoFeed?.SetValue(BoolType.True);
                source.Capabilities.ICapXResolution?.SetValue(400);
                source.Capabilities.ICapYResolution?.SetValue(400);

                _session.DataTransferred += (s, args) =>
                {
                    if (args.NativeData != null)
                    {
                        using var stream = args.GetNativeImageStream();
                        if (stream != null)
                        {
                            string path = System.IO.Path.Combine(_saveDir, $"Page{++_imageCount}.bmp");
                            using var bmp = new Bitmap(stream);
                            bmp.Save(path);
                        }
                    }
                };

                _session.TransferError += (s, args) =>
                {
                    MessageBox.Show("Transfer error: " + args.Exception?.Message);
                };

                source.Enable(SourceEnableMode.NoUI, false, IntPtr.Zero);

                Console.WriteLine("Session State: ", _session.State);

                while (_session.State == 5) Thread.Sleep(10000);
                source.Close();
                _session.Close();
            });
        }

        private async Task InitScannerAsync()
        {
            AddStatus("Initializing Twain session...");

            var appId = TWIdentity.CreateFromAssembly(DataGroups.Image, typeof(Form1).Assembly);
            _session = new TwainSession(appId);

            _session.StateChanged += (s, ev) =>
            {
                AddStatus($"TWAIN session state changed.");
            };
            _session.DataTransferred += (s, ev) =>
            {
                if (ev.NativeData != null)
                {
                    string path = System.IO.Path.Combine(Environment.CurrentDirectory, $"ID_Page{++_imageCount}.bmp");
                    using var stream = ev.GetNativeImageStream();
                    if (stream != null)
                    {
                        using var bmp = new Bitmap(stream);
                        bmp.Save(path);
                        AddStatus($"Saved: {path}");

                        if (_imageCount == 1)
                        {
                            Invoke(() =>
                            {
                                picPreview.Image?.Dispose();
                                picPreview.Image = new Bitmap(path);
                            });
                        }
                    }
                    else
                    {
                        AddStatus("Failed to retrieve image stream.");
                    }
                }
            };

            _session.TransferError += (s, ev) =>
            {
                AddStatus("Transfer error: " + ev.Exception?.Message);
            };

            //_session.SourceChanged += (s, ev) =>
            //{
            //    if (ev.Change == SourceChange.Added)
            //    {
            //        AddStatus($"Scanner connected: {ev.Source?.Name ?? "(Unknown)"}");

            //        _ = Task.Run(() =>
            //        {
            //            // Small delay to let the device settle
            //            Thread.Sleep(1500);
            //            StartScanWorkflow();
            //        });
            //    }
            //};

            _session.SourceChanged += (s, ev) =>
            {
                Thread.Sleep(1500);
                StartScanWorkflow();
            };
            StartScanWorkflow();

            _session.Open(); // non-blocking in this context

            AddStatus("Waiting for scanner to be connected...");
        }

        private void StartScanWorkflow()
        {
            try
            {
                //if (_session.State < 3) return;

                var source = _session.FirstOrDefault();
                if (source == null)
                {
                    AddStatus("No scanner found.");
                    return;
                }

                AddStatus("Opening source...");
                source.Open();

                // Set capabilities
                source.Capabilities.CapDuplexEnabled?.SetValue(BoolType.True);
                source.Capabilities.CapFeederEnabled?.SetValue(BoolType.True);
                source.Capabilities.CapAutoFeed?.SetValue(BoolType.True);
                source.Capabilities.ICapXResolution?.SetValue(300);
                source.Capabilities.ICapYResolution?.SetValue(300);

                //if (!Directory.Exists(saveDir))
                //    Directory.CreateDirectory(saveDir);

                _imageCount = 0;

                AddStatus("Enabling source for scanning...");
                source.Enable(SourceEnableMode.NoUI, false, IntPtr.Zero);

                AddStatus("Scan started...");
            }
            catch (Exception ex)
            {
                AddStatus("Scan error: " + ex.Message);
            }
        }

        private void AddStatus(string message)
        {
            Invoke(() => lstStatus.Items.Add($"{DateTime.Now:T} - {message}"));
        }
    }
}
