using System;
using System.Diagnostics;
using System.Security.Principal;
using System.ServiceProcess;

namespace NexusAgent
{
    static class Program
    {
        public const string ServiceName = "NEXUS Agent";
        public const string ServiceDisplayName = "NEXUS Agent";
        public const string ServiceDescription = "NEXUS Windows Agent for Active Directory integration and PowerShell command execution.";

        static void Main(string[] args)
        {
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
                    RunConsoleMode();
                    return;
                }
            }

            if (Environment.UserInteractive)
            {
                if (IsServiceInstalled())
                {
                    Console.WriteLine("NEXUS Agent is already installed.");
                    Console.WriteLine("Starting the service...");
                    StartService();
                    Console.WriteLine();
                    Console.WriteLine("Press any key to exit...");
                    Console.ReadKey();
                    return;
                }

                if (!IsAdministrator())
                {
                    RelaunchAsAdmin("--install");
                    return;
                }

                InstallService();
                return;
            }

            ServiceBase[] servicesToRun = new ServiceBase[] { new NexusAgentService() };
            ServiceBase.Run(servicesToRun);
        }

        private static bool IsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        private static void RelaunchAsAdmin(string arguments)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = System.Reflection.Assembly.GetExecutingAssembly().Location;
            startInfo.Arguments = arguments;
            startInfo.UseShellExecute = true;
            startInfo.Verb = "runas";

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
            ServiceController sc = null;
            try
            {
                sc = new ServiceController(ServiceName);
                ServiceControllerStatus status = sc.Status;
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (sc != null)
                {
                    sc.Dispose();
                }
            }
        }

        private static void InstallService()
        {
            Console.WriteLine("========================================");
            Console.WriteLine("       NEXUS Agent Installer            ");
            Console.WriteLine("========================================");
            Console.WriteLine();

            if (!IsAdministrator())
            {
                Console.WriteLine("Error: Administrator privileges required.");
                Console.WriteLine("Please run as Administrator.");
                Console.WriteLine();
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                return;
            }

            string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            
            if (string.IsNullOrEmpty(exePath))
            {
                Console.WriteLine("Error: Could not determine executable path.");
                Console.WriteLine();
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                return;
            }

            try
            {
                Console.WriteLine("Installing NEXUS Agent service...");
                
                ProcessStartInfo scCreate = new ProcessStartInfo();
                scCreate.FileName = "sc.exe";
                scCreate.Arguments = string.Format("create \"{0}\" binPath= \"\\\"{1}\\\"\" start= auto DisplayName= \"{2}\"", ServiceName, exePath, ServiceDisplayName);
                scCreate.UseShellExecute = false;
                scCreate.RedirectStandardOutput = true;
                scCreate.RedirectStandardError = true;
                scCreate.CreateNoWindow = true;

                Process process = Process.Start(scCreate);
                process.WaitForExit();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();

                if (process.ExitCode != 0 && !output.Contains("exists"))
                {
                    Console.WriteLine("Warning: " + error);
                }

                ProcessStartInfo scDesc = new ProcessStartInfo();
                scDesc.FileName = "sc.exe";
                scDesc.Arguments = string.Format("description \"{0}\" \"{1}\"", ServiceName, ServiceDescription);
                scDesc.UseShellExecute = false;
                scDesc.CreateNoWindow = true;
                Process descProcess = Process.Start(scDesc);
                descProcess.WaitForExit();

                ProcessStartInfo scFailure = new ProcessStartInfo();
                scFailure.FileName = "sc.exe";
                scFailure.Arguments = string.Format("failure \"{0}\" reset= 86400 actions= restart/60000/restart/60000/restart/60000", ServiceName);
                scFailure.UseShellExecute = false;
                scFailure.CreateNoWindow = true;
                Process failProcess = Process.Start(scFailure);
                failProcess.WaitForExit();

                Console.WriteLine("[OK] Service installed successfully!");
                Console.WriteLine();

                Console.WriteLine("Starting NEXUS Agent service...");
                StartService();
                
                Console.WriteLine();
                Console.WriteLine("========================================");
                Console.WriteLine("    Installation Complete!              ");
                Console.WriteLine("                                        ");
                Console.WriteLine("    The NEXUS Agent is now running      ");
                Console.WriteLine("    and will start automatically        ");
                Console.WriteLine("    when Windows starts.                ");
                Console.WriteLine("========================================");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error installing service: " + ex.Message);
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static void UninstallService()
        {
            Console.WriteLine("========================================");
            Console.WriteLine("       NEXUS Agent Uninstaller          ");
            Console.WriteLine("========================================");
            Console.WriteLine();

            if (!IsAdministrator())
            {
                RelaunchAsAdmin("--uninstall");
                return;
            }

            try
            {
                Console.WriteLine("Stopping NEXUS Agent service...");
                ServiceController sc = null;
                try
                {
                    sc = new ServiceController(ServiceName);
                    if (sc.Status != ServiceControllerStatus.Stopped)
                    {
                        sc.Stop();
                        sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30));
                    }
                    Console.WriteLine("[OK] Service stopped.");
                }
                catch
                {
                    Console.WriteLine("Service was not running.");
                }
                finally
                {
                    if (sc != null)
                    {
                        sc.Dispose();
                    }
                }

                Console.WriteLine("Removing NEXUS Agent service...");
                ProcessStartInfo scDelete = new ProcessStartInfo();
                scDelete.FileName = "sc.exe";
                scDelete.Arguments = string.Format("delete \"{0}\"", ServiceName);
                scDelete.UseShellExecute = false;
                scDelete.RedirectStandardOutput = true;
                scDelete.CreateNoWindow = true;

                Process process = Process.Start(scDelete);
                process.WaitForExit();

                Console.WriteLine("[OK] Service removed successfully!");
                Console.WriteLine();
                Console.WriteLine("========================================");
                Console.WriteLine("    Uninstallation Complete!            ");
                Console.WriteLine("========================================");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error uninstalling service: " + ex.Message);
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static void StartService()
        {
            ServiceController sc = null;
            try
            {
                sc = new ServiceController(ServiceName);
                if (sc.Status != ServiceControllerStatus.Running)
                {
                    sc.Start();
                    sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
                }
                Console.WriteLine("[OK] Service started successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error starting service: " + ex.Message);
            }
            finally
            {
                if (sc != null)
                {
                    sc.Dispose();
                }
            }
        }

        private static void RunConsoleMode()
        {
            Console.WriteLine("Running NEXUS Agent in console mode...");
            Console.WriteLine("Press Ctrl+C to stop.");
            
            NexusAgentService service = new NexusAgentService();
            service.StartConsoleMode();
            
            while (true)
            {
                System.Threading.Thread.Sleep(1000);
            }
        }
    }
}
