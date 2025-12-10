using System;
using System.ServiceProcess;

namespace NexusAgent
{
    static class Program
    {
        static void Main()
        {
            if (OperatingSystem.IsWindows())
            {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                    new NexusAgentService()
                };
                ServiceBase.Run(ServicesToRun);
            }
            else
            {
                Console.WriteLine("This agent only runs on Windows");
            }
        }
    }
}
