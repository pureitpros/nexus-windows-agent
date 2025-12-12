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
            string exePath = Assembly.GetExecutingAssembly().Location;

            if (args.Length > 0)
            {
                string command = args[0].ToLower();

                if (command == "/install" || command == "-install" || command == "--install")
                {
                    Console.WriteLine("Installing NEXUS Agent service...");
                    try
                    {
                        ManagedInstallerClass.InstallHelper(new string[] { exePath });
                        Console.WriteLine("Service installed successfully.");
                        Console.WriteLine("Starting the service...");
                        
                        try
                        {
                            using (ServiceController sc = new ServiceController("NexusAgent"))
                            {
                                sc.Start();
                                sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
                                Console.WriteLine("Service started successfully.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error starting service: " + ex.Message);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error installing service: " + ex.Message);
                    }
                    
                    Console.WriteLine("\nPress any key to exit...");
                    Console.ReadKey();
                    return;
                }
                else if (command == "/uninstall" || command == "-uninstall" || command == "--uninstall")
                {
                    Console.WriteLine("Uninstalling NEXUS Agent service...");
                    try
                    {
                        // Stop the service first if running
                        try
                        {
                            using (ServiceController sc = new ServiceController("NexusAgent"))
                            {
                                if (sc.Status == ServiceControllerStatus.Running)
                                {
                                    Console.WriteLine("Stopping service...");
                                    sc.Stop();
                                    sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30));
                                }
                            }
                        }
                        catch { }
                        
                        ManagedInstallerClass.InstallHelper(new string[] { "/u", exePath });
                        Console.WriteLine("Service uninstalled successfully.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error uninstalling service: " + ex.Message);
                    }
                    
                    Console.WriteLine("\nPress any key to exit...");
                    Console.ReadKey();
                    return;
                }
            }

            // Check if already installed
            bool serviceExists = false;
            try
            {
                using (ServiceController sc = new ServiceController("NexusAgent"))
                {
                    var status = sc.Status;
                    serviceExists = true;
                }
            }
            catch { }

            if (!serviceExists)
            {
                // Not installed and no args - install automatically
                Console.WriteLine("NEXUS Agent is not installed. Installing...");
                try
                {
                    ManagedInstallerClass.InstallHelper(new string[] { exePath });
                    Console.WriteLine("Service installed successfully.");
                    Console.WriteLine("Starting the service...");
                    
                    try
                    {
                        using (ServiceController sc = new ServiceController("NexusAgent"))
                        {
                            sc.Start();
                            sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
                            Console.WriteLine("Service started successfully.");
                            Console.WriteLine("\nThe NEXUS Agent is now running as a Windows service.");
                            Console.WriteLine("You can close this window.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error starting service: " + ex.Message);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error installing service: " + ex.Message);
                }
                
                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("NEXUS Agent is already installed.");
                Console.WriteLine("Starting the service...");
                try
                {
                    using (ServiceController sc = new ServiceController("NexusAgent"))
                    {
                        if (sc.Status != ServiceControllerStatus.Running)
                        {
                            sc.Start();
                            sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
                        }
                        Console.WriteLine("Service is running.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error starting service: " + ex.Message);
                }
                
                Console.WriteLine("\nPress any key to exit...");
                Console.ReadKey();
            }
        }
    }
}
