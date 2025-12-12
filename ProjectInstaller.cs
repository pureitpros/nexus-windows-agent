using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;

namespace NexusAgent
{
    [RunInstaller(true)]
    public class ProjectInstaller : Installer
    {
        private ServiceProcessInstaller processInstaller;
        private ServiceInstaller serviceInstaller;

        public ProjectInstaller()
        {
            processInstaller = new ServiceProcessInstaller();
            serviceInstaller = new ServiceInstaller();

            // Run as Local System account
            processInstaller.Account = ServiceAccount.LocalSystem;

            // Service configuration
            serviceInstaller.ServiceName = "NexusAgent";
            serviceInstaller.DisplayName = "NEXUS Agent";
            serviceInstaller.Description = "NEXUS on-premises agent for Active Directory integration";
            serviceInstaller.StartType = ServiceStartMode.Automatic;

            Installers.Add(processInstaller);
            Installers.Add(serviceInstaller);
        }
    }
}
