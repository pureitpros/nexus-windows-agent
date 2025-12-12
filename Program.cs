using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.ServiceProcess;

namespace NexusAgent
{
    public class Program
    {
        public const string ServiceName = "NEXUS Agent";
        public const string ServiceDisplayName = "NEXUS Agent";
        public const string ServiceDescription = "NEXUS Windows Agent for Active Directory integration and PowerShell command execution.";

        public static void Main(string[] args)
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                Console.WriteLine("This application only runs on Windows.");
                return;
            }

            // Check for command line arguments
            if (args.Length > 0)
            {
                string command = args[0].ToLower();
                
                if (command == "--install" || command == "-i")
                {
                    InstallService();
                    return;
                }
                else if (command == "--uninstall" || command == "-u")
                {
                    UninstallService();
                    return;
                }
                else if (command == "--console" || command == "-c")
                {
                    // Run in console mode for debugging
                    RunConsoleMode();
                    return;
                }
            }

            // If running interactively (double-clicked), install the service
            if (Environment.UserInteractive)
            {
                // Check if already installed
                if (IsServiceInstalled())
                {
                    Console.WriteLine("NEXUS Agent is already installed.");
                    Console.WriteLine("Starting the service...");
                    StartService();
                    Console.WriteLine("\nPress any key to exit...");
                    Console.ReadKey();
                    return;
                }

                // Need to install - check for admin rights
                if (!IsAdministrator())
                {
                    // Relaunch with admin rights
                    RelaunchAsAdmin("--install");
                    return;
                }

                InstallService();
                return;
            }

            // Running as a service
            ServiceBase.Run(new NexusAgentService());
        }

        private static bool IsAdministrator()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        private static void RelaunchAsAdmin(string arguments)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = Process.GetCurrentProcess().MainModule?.FileName ?? "NexusAgent.exe",
                Arguments = arguments,
                UseShellExecute = true,
                Verb = "runas" // This triggers UAC prompt
            };

            try
            {
                Process.Start(startInfo);
            }
            catch (System.ComponentModel.Win32Exception)
            {
                Console.WriteLine("Administrator privileges are required to install the service.");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
            }
        }

        private static bool IsServiceInstalled()
        {
            try
            {
                using (ServiceController sc = new ServiceController(ServiceName))
                {
                    var status = sc.Status;
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private static void InstallService()
        {
            Console.WriteLine("╔════════════════════════════════════════╗");
            Console.WriteLine("║       NEXUS Agent Installer            ║");
            Console.WriteLine("╚════════════════════════════════════════╝");
            Console.WriteLine();

            if (!IsAdministrator())
            {
                Console.WriteLine("Error: Administrator privileges required.");
                Console.WriteLine("Please run as Administrator.");
                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
                return;
            }

            string exePath = Process.GetCurrentProcess().MainModule?.FileName ?? "";
            
            if (string.IsNullOrEmpty(exePath))
            {
                Console.WriteLine("Error: Could not determine executable path.");
                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
                return;
            }

            try
            {
                Console.WriteLine("Installing NEXUS Agent service...");
                
                // Create the service using sc.exe
                ProcessStartInfo scCreate = new ProcessStartInfo
                {
                    FileName = "sc.exe",
                    Arguments = $"create \"{ServiceName}\" binPath= \"\\\"{exePath}\\\"\" start= auto DisplayName= \"{ServiceDisplayName}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(scCreate)!)
                {
                    process.WaitForExit();
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    if (process.ExitCode != 0 && !output.Contains("exists"))
                    {
                        Console.WriteLine($"Warning: {error}");
                    }
                }

                // Set service description
                ProcessStartInfo scDesc = new ProcessStartInfo
                {
                    FileName = "sc.exe",
                    Arguments = $"description \"{ServiceName}\" \"{ServiceDescription}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                Process.Start(scDesc)?.WaitForExit();

                // Configure service recovery (restart on failure)
                ProcessStartInfo scFailure = new ProcessStartInfo
                {
                    FileName = "sc.exe",
                    Arguments = $"failure \"{ServiceName}\" reset= 86400 actions= restart/60000/restart/60000/restart/60000",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                Process.Start(scFailure)?.WaitForExit();

                Console.WriteLine("✓ Service installed successfully!");
                Console.WriteLine();

                // Start the service
                Console.WriteLine("Starting NEXUS Agent service...");
                StartService();
                
                Console.WriteLine();
                Console.WriteLine("╔════════════════════════════════════════╗");
                Console.WriteLine("║    Installation Complete!              ║");
                Console.WriteLine("║                                        ║");
                Console.WriteLine("║    The NEXUS Agent is now running      ║");
                Console.WriteLine("║    and will start automatically        ║");
                Console.WriteLine("║    when Windows starts.                ║");
                Console.WriteLine("╚════════════════════════════════════════╝");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error installing service: {ex.Message}");
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }

        private static void UninstallService()
        {
            Console.WriteLine("╔════════════════════════════════════════╗");
            Console.WriteLine("║       NEXUS Agent Uninstaller          ║");
            Console.WriteLine("╚════════════════════════════════════════╝");
            Console.WriteLine();

            if (!IsAdministrator())
            {
                RelaunchAsAdmin("--uninstall");
                return;
            }

            try
            {
                // Stop the service first
                Console.WriteLine("Stopping NEXUS Agent service...");
                try
                {
                    using (ServiceController sc = new ServiceController(ServiceName))
                    {
                        if (sc.Status != ServiceControllerStatus.Stopped)
                        {
                            sc.Stop();
                            sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30));
                        }
                    }
                    Console.WriteLine("✓ Service stopped.");
                }
                catch
                {
                    Console.WriteLine("Service was not running.");
                }

                // Delete the service
                Console.WriteLine("Removing NEXUS Agent service...");
                ProcessStartInfo scDelete = new ProcessStartInfo
                {
                    FileName = "sc.exe",
                    Arguments = $"delete \"{ServiceName}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(scDelete)!)
                {
                    process.WaitForExit();
                }

                Console.WriteLine("✓ Service removed successfully!");
                Console.WriteLine();
                Console.WriteLine("╔════════════════════════════════════════╗");
                Console.WriteLine("║    Uninstallation Complete!            ║");
                Console.WriteLine("╚════════════════════════════════════════╝");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error uninstalling service: {ex.Message}");
            }

            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }

        private static void StartService()
        {
            try
            {
                using (ServiceController sc = new ServiceController(ServiceName))
                {
                    if (sc.Status != ServiceControllerStatus.Running)
                    {
                        sc.Start();
                        sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
                    }
                    Console.WriteLine("✓ Service started successfully!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error starting service: {ex.Message}");
            }
        }

        private static void RunConsoleMode()
        {
            Console.WriteLine("Running NEXUS Agent in console mode...");
            Console.WriteLine("Press Ctrl+C to stop.");
            
            var service = new NexusAgentService();
            service.StartConsoleMode();
            
            // Keep running until Ctrl+C
            var exitEvent = new System.Threading.ManualResetEvent(false);
            Console.CancelKeyPress += (sender, e) =>
            {
                e.Cancel = true;
                exitEvent.Set();
            };
            exitEvent.WaitOne();
            
            service.StopConsoleMode();
        }
    }
}
