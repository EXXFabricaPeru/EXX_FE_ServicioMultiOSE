using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace ServicioConectorFE
{
    static class Program
    {
        static void Main()
        {
            if (Environment.UserInteractive)
            {
                try
                {
                    Core.Main.IniciarApp();
                }
                catch (Exception ex)
                {
                    throw;
                }
                finally
                {
                    GC.Collect();
                }
            }
            else
            {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                new ServicioConector()
                };
                ServiceBase.Run(ServicesToRun);
            }
        }
    }
}
