using System.ServiceProcess;

namespace opc_bridge
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new OpcServerService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
