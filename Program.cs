using System;
using System.ServiceProcess;
using System.Configuration.Install;
using System.Reflection;

namespace NexusAgent
{
    static class Program
    {
        static void Main(string[] args)
        {
            // Get the path to this executable
            string exePath = Assembly.GetExecutingAssembly().Location;
            
            // Check if running as a service or in console mode
            if (Environment.UserInteractive)
            {
                // Running from command line - handle install/uninstall
                if (args.Length > 0)
                {
                    switch (args[0].ToLower())
                    {
                        case "/install":
                        case "-install":
                        case "--install":
                            ManagedInstallerClass.InstallHelper(new string[] { exePath });
                            Console.WriteLine("NEXUS Agent service installed successfully.");
                            return;
                        case "/uninstall":
                        case "-uninstall":
                        case "--uninstall":
                            ManagedInstallerClass.InstallHelper(new string[] { "/u", exePath });
                            Console.WriteLine("NEXUS Agent service uninstalled successfully.");
                            return;
                    }
                }

                // Default: Install and start the service
                try
                {
                    Console.WriteLine("Installing NEXUS Agent service...");
                    
                    // Check if already installed
                    ServiceController[] services = ServiceController.GetServices();
                    bool isInstalled = false;
                    foreach (ServiceController svc in services)
                    {
                        if (svc.ServiceName == "NexusAgent")
                        {
                            isInstalled = true;
                            break;
                        }
                    }

                    if (!isInstalled)
                    {
                        ManagedInstallerClass.InstallHelper(new string[] { exePath });
                        Console.WriteLine("Service installed.");
                    }
                    else
                    {
                        Console.WriteLine("NEXUS Agent is already installed.");
                    }

                    // Start the service
                    Console.WriteLine("Starting the service...");
                    using (ServiceController sc = new ServiceController("NexusAgent"))
                    {
                        if (sc.Status != ServiceControllerStatus.Running)
                        {
                            sc.Start();
                            sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
                        }
                        Console.WriteLine("NEXUS Agent service is now running.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }

                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
            }
            else
            {
                // Running as Windows Service
                ServiceBase[] ServicesToRun = new ServiceBase[] { new NexusAgentService() };
                ServiceBase.Run(ServicesToRun);
            }
        }
    }
}
