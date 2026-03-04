using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace XlsxToKompasTblLauncher
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            try
            {
                Process currentProcess = Process.GetCurrentProcess();
                string exePath = Application.ExecutablePath;
                if (currentProcess != null && currentProcess.MainModule != null && !string.IsNullOrWhiteSpace(currentProcess.MainModule.FileName))
                {
                    exePath = currentProcess.MainModule.FileName;
                }
                string exeDir = Path.GetDirectoryName(exePath) ?? AppDomain.CurrentDomain.BaseDirectory;
                string appRoot = Path.GetFullPath(Path.Combine(exeDir, ".."));
                string guiScript = Path.Combine(appRoot, "scripts", "gui_import.ps1");

                if (!File.Exists(guiScript))
                {
                    MessageBox.Show(
                        "gui_import.ps1 not found: " + guiScript,
                        "xlsx-to-kompas-tbl app",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                    return 2;
                }

                string psExe = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.System),
                    "WindowsPowerShell",
                    "v1.0",
                    "powershell.exe"
                );

                if (!File.Exists(psExe))
                {
                    psExe = "powershell.exe";
                }

                string arguments =
                    "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File \"" +
                    guiScript +
                    "\"";

                var startInfo = new ProcessStartInfo
                {
                    FileName = psExe,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                Process.Start(startInfo);
                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Failed to start GUI: " + ex.Message,
                    "xlsx-to-kompas-tbl app",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return 1;
            }
        }
    }
}
