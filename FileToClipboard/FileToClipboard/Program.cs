using FileToClipboard.Properties;
using System;
using System.Collections.Specialized;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FileToClipboard
{
    static class Program
    {
        static string path = Environment.ExpandEnvironmentVariables(Settings.Default.WatchPath);
        static DirectoryInfo dir = new DirectoryInfo(path);
        static NotifyIcon tray = null;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            dir.Create();
            dir.Clear();

            var watcher = new FileSystemWatcher(path);
            watcher.Created += OnFileCreated;
            watcher.EnableRaisingEvents = true;

            Application.ApplicationExit += Application_ApplicationExit;
            tray = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                BalloonTipTitle = "FileToClipboard",
                BalloonTipIcon = ToolTipIcon.Info,
                Text = "FileToClipboard",
                Visible = true,
                ContextMenu = new ContextMenu(new MenuItem[]
                {
                    new MenuItem("Exit", (sender, e) => {
                        Application.Exit();
                    })
                })
            };

            Application.Run();
        }

        static void Application_ApplicationExit(object sender, EventArgs e)
        {
            tray.Visible = false;
        }

        static async void OnFileCreated(object sender, FileSystemEventArgs e)
        {
            var file = new FileInfo(e.FullPath);
            while (IsFileLocked(file))
            {
                await Task.Delay(500);
            }

            // STA is needed because System.Windows.Forms.Clipboard is an OLE thing
            STARun(() =>
            {
                var dotIndex = e.Name.LastIndexOf('.');
                if (dotIndex > 0)
                {
                    switch (e.Name.Substring(dotIndex).ToLower())
                    {
                    case ".png":
                        {
                            try
                            {
                                using (var bmp = new Bitmap(e.FullPath))
                                using (var pngStream = new MemoryStream(File.ReadAllBytes(e.FullPath)))
                                {
                                    var data = new DataObject();
                                    data.SetImage(bmp);
                                    data.SetData("PNG", false, pngStream);
                                    Clipboard.SetDataObject(data, true);
                                    Notify(e.Name, "PNG", new
                                    {
                                        Resolution = $"{bmp.Width}×{bmp.Height}"
                                    });
                                }
                            }
                            catch (ArgumentException ex)
                            {
                                Notify(e.Name, "PNG", new
                                {
                                    Error = ex.ToString()
                                });
                            }
                            break;
                        }
                    case ".jpg":
                    case ".gif":
                    case ".bmp":
                    case ".jpeg":
                        {
                            try
                            {
                                using (var bmp = new Bitmap(e.FullPath))
                                {
                                    Clipboard.SetImage(bmp);
                                    Notify(e.Name, "Bitmap", new
                                    {
                                        Resolution = $"{bmp.Width}×{bmp.Height}"
                                    });
                                }
                            }
                            catch (ArgumentException ex)
                            {
                                Notify(e.Name, "Bitmap", new
                                {
                                    Error = ex.ToString()
                                }, ToolTipIcon.Error);
                            }
                            break;
                        }
                    #region Text File
                    // Plain Text
                    case ".txt":
                    case ".log":
                    // Data Serialization
                    case ".json":
                    case ".xml":
                    // Config
                    case ".ini":
                    case ".conf":
                    case ".config":
                    case ".reg":
                    // Shell Script
                    case ".bat":
                    case ".cmd":
                    case ".ps1":
                    case ".sh":
                    // .NET
                    case ".cs":
                    case ".vb":
                    case ".xaml":
                    case ".cshtml":
                    case ".vbhtml":
                    // Web - Essential
                    case ".html":
                    case ".htm":
                    case ".js":
                    case ".css":
                    case ".svg":
                    // Web - More
                    case ".ts":
                    case ".jsx":
                    case ".tsx":
                    case ".less":
                    case ".sass":
                    // C/C++
                    case ".c":
                    case ".cpp":
                    case ".h":
                    // Other
                    case ".php":
                    case ".py":
                    case ".java":
                    case ".jsp":
                    case ".pl":
                    case ".ahk":
                    #endregion
                        {
                            Clipboard.SetText(File.ReadAllText(e.FullPath));
                            Notify(e.Name, "Text", new { });
                            break;
                        }
                    default:
                        {
                            Notify(e.Name, "Unknown", new
                            {
                                Error = "Unknown file type."
                            }, ToolTipIcon.Error);
                            break;
                        }
                    }
                }
                else
                {
                    Notify(e.Name, "Unknown", new
                    {
                        Error = "Unknown file type."
                    }, ToolTipIcon.Error);
                }

                Task.Delay(1000).ContinueWith(t =>
                {
                    if (File.Exists(e.FullPath))
                    {
                        File.Delete(e.FullPath);
                    }
                    else if (Directory.Exists(e.FullPath))
                    {
                        Directory.Delete(e.FullPath, true);
                    }
                });
            });
        }

        static void Notify(string filename, string type, object custom, ToolTipIcon icon = ToolTipIcon.Info)
        {
            var text = $"File: {filename}\nType: {type}";
            custom.GetType().GetProperties()
                .Where(p => p.PropertyType == typeof(string))
                .ToList()
                .ForEach(p => text += $"\n{p.Name}: {(string)p.GetValue(custom)}");
            tray.BalloonTipText = text;
            tray.BalloonTipIcon = icon;
            tray.ShowBalloonTip(5000);
        }

        static void STARun(ThreadStart func)
        {
            var t = new Thread(func);
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        // https://stackoverflow.com/a/937558/8048600
        static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }
    }

    public static class DirectoryInfoExtension
    {
        public static void Clear(this DirectoryInfo dir)
        {
            dir.EnumerateFiles().ToList().ForEach(f => f.Delete());
            dir.EnumerateDirectories().ToList().ForEach(d => d.Delete(true));
        }
    }
}
