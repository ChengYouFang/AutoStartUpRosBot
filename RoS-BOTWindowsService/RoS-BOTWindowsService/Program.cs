using System.ServiceProcess;

namespace RoS_BOTWindowsService
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new Ros_Bot_Service()
            };

            ServiceBase.Run(ServicesToRun);
        }
    }
}
