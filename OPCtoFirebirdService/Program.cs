using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace OPCtoFirebirdService
{
    static class Program
    {
        public static OPCtoFirebirdService MyService;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
            { 
                MyService = new OPCtoFirebirdService() 
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
