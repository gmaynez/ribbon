using System;
using System.Diagnostics;
using System.IO;

namespace Ribbon.Vsto
{
    internal static class BrokerLocator
    {
        public static void StartBroker()
        {
            var broker = LocateBroker();
            var startInfo = new ProcessStartInfo
            {
                FileName = broker,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                WorkingDirectory = Path.GetDirectoryName(broker)
            };
            Process.Start(startInfo);
        }

        private static string LocateBroker()
        {
            var configured = Environment.GetEnvironmentVariable("RIBBON_BROKER_PATH");
            if (!string.IsNullOrWhiteSpace(configured) && File.Exists(configured))
            {
                return Path.GetFullPath(configured);
            }

            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            var deployed = Path.Combine(baseDirectory, "Ribbon.Broker.exe");
            if (File.Exists(deployed))
            {
                return deployed;
            }

            var cursor = new DirectoryInfo(baseDirectory);
            for (var depth = 0; depth < 8 && cursor != null; depth++, cursor = cursor.Parent)
            {
                foreach (var configuration in new[] { "Debug", "Release" })
                {
                    var candidate = Path.Combine(cursor.FullName, "Ribbon.Broker", "bin", configuration, "net10.0-windows", "Ribbon.Broker.exe");
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }

            throw new FileNotFoundException(
                "Ribbon.Broker.exe was not found. Build Ribbon.Broker or set RIBBON_BROKER_PATH to its full path.");
        }
    }
}
