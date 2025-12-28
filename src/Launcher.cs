using System;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using System.Security.Principal;
using System.Runtime.InteropServices;
using System.Threading;

namespace KMSActivator
{
    static class Program
    {
        // P/Invoke for icon injection
        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        
        [DllImport("user32.dll")]
        static extern IntPtr LoadImage(IntPtr hInst, string lpszName, uint uType, int cxDesired, int cyDesired, uint fuLoad);
        
        const int SW_HIDE = 0;
        const uint WM_SETICON = 0x0080;
        const uint ICON_SMALL = 0;
        const uint ICON_BIG = 1;
        const uint IMAGE_ICON = 1;
        const uint LR_LOADFROMFILE = 0x0010;
        
        private static Process psProcess;
        private static string iconPath = "";
        
        [STAThread]
        static void Main()
        {
            if (!IsAdministrator())
            {
                RestartAsAdmin();
                return;
            }

            try
            {
                string tempDir = Path.Combine(Path.GetTempPath(), "KMS_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tempDir);

                ExtractEmbeddedZip(tempDir);

                iconPath = Path.Combine(tempDir, "app.ico");
                ExtractIcon(iconPath);

                string mainScript = Path.Combine(tempDir, "KMS_Activator_GUI.ps1");
                
                if (!File.Exists(mainScript))
                {
                    MessageBox.Show("Failed to extract application files", "Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                Application.EnableVisualStyles();
                
                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = string.Format("-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"{0}\"", mainScript),
                    UseShellExecute = false,
                    WorkingDirectory = tempDir,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                psProcess = Process.Start(psi);
                
                // Monitor PowerShell and inject icon
                var monitorThread = new Thread(() =>
                {
                    bool iconSet = false;
                    
                    while (!psProcess.HasExited)
                    {
                        try
                        {
                            psProcess.Refresh();
                            if (psProcess.MainWindowHandle != IntPtr.Zero && !iconSet && File.Exists(iconPath))
                            {
                                IntPtr hIcon = LoadImage(IntPtr.Zero, iconPath, IMAGE_ICON, 0, 0, LR_LOADFROMFILE);
                                if (hIcon != IntPtr.Zero)
                                {
                                    SendMessage(psProcess.MainWindowHandle, WM_SETICON, new IntPtr(ICON_SMALL), hIcon);
                                    SendMessage(psProcess.MainWindowHandle, WM_SETICON, new IntPtr(ICON_BIG), hIcon);
                                    iconSet = true;
                                }
                            }
                            
                            // Hide PowerShell console windows
                            var processes = Process.GetProcessesByName("powershell");
                            foreach (var p in processes)
                            {
                                if (p.MainWindowHandle != IntPtr.Zero && string.IsNullOrEmpty(p.MainWindowTitle))
                                {
                                    ShowWindow(p.MainWindowHandle, SW_HIDE);
                                }
                            }
                        }
                        catch { }
                        
                        Thread.Sleep(500);
                    }
                    
                    // Cleanup
                    try
                    {
                        Directory.Delete(tempDir, true);
                    }
                    catch { }
                });
                
                monitorThread.IsBackground = true;
                monitorThread.Start();
                
                // Wait for PowerShell to finish
                psProcess.WaitForExit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Error: {0}", ex.Message), "KMS Activator", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        static void ExtractEmbeddedZip(string targetDir)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("KMSActivator.payload.zip"))
            {
                if (stream == null) 
                    throw new Exception("Embedded ZIP not found");
                    
                using (var zip = new ZipArchive(stream, ZipArchiveMode.Read))
                {
                    zip.ExtractToDirectory(targetDir);
                }
            }
        }

        static void ExtractIcon(string targetPath)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var iconStream = assembly.GetManifestResourceStream("KMSActivator.app.ico");
                
                if (iconStream != null)
                {
                    using (var fileStream = File.Create(targetPath))
                    {
                        iconStream.CopyTo(fileStream);
                    }
                }
            }
            catch { }
        }

        static bool IsAdministrator()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        static void RestartAsAdmin()
        {
            var exeName = Assembly.GetExecutingAssembly().Location;
            var psi = new ProcessStartInfo
            {
                FileName = exeName,
                Verb = "runas",
                UseShellExecute = true
            };
            
            try { Process.Start(psi); } catch { }
        }
    }
}
