using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace ChatFormatterPro
{
    public static class FileOpener
    {
        public static void Open(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath))
                    return;

                // ✅ wait until file really exists and is ready (important for PDF/Excel)
                for (int i = 0; i < 20; i++)
                {
                    if (File.Exists(filePath))
                    {
                        // try to open file stream to check lock
                        try
                        {
                            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                            break; // file is readable
                        }
                        catch
                        {
                            // still locked -> wait
                        }
                    }

                    Thread.Sleep(150);
                }

                if (!File.Exists(filePath))
                    return;

                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            catch
            {
                // fallback: open folder if file cannot be opened
                try
                {
                    var folder = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder))
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = folder,
                            UseShellExecute = true
                        });
                    }
                }
                catch { }
            }
        }
    }
}
