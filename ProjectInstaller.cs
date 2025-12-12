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
            processInstaller.Account = ServiceAccount.LocalSystem;

            serviceInstaller = new ServiceInstaller();
            serviceInstaller.ServiceName = "NexusAgent";
            serviceInstaller.DisplayName = "NEXUS Agent";
            serviceInstaller.Description = "NEXUS Windows Agent for Active Directory integration and PowerShell execution";
            serviceInstaller.StartType = ServiceStartMode.Automatic;

            Installers.Add(processInstaller);
            Installers.Add(serviceInstaller);
        }
    }
}
