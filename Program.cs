using System.ServiceProcess;

namespace NexusAgent
{
    static class Program
    {
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new NexusAgentService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
