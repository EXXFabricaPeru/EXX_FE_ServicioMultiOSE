using ServicioConectorFE.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ServicioConectorFE
{
    public partial class ServicioConector: ServiceBase
    {
        //public Thread workerThread;

        private CancellationTokenSource cancellationTokenSource;
        private Task workerTask;

        public ServicioConector()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            cancellationTokenSource = new CancellationTokenSource();
            var token = cancellationTokenSource.Token;

            workerTask = Task.Run(() => DoWorkAsync(token), token);

            //workerThread = new Thread(DoWork);
            //workerThread.Start();
        }

        protected override void OnStop()
        {
            Globals.Log("Servicio detenido manualmente por el usuario.");

            if (cancellationTokenSource != null)
            {
                cancellationTokenSource.Cancel();

                try
                {
                    workerTask?.Wait(); // Espera que termine bien
                }
                catch (AggregateException ae)
                {
                    foreach (var ex in ae.InnerExceptions)
                        Globals.Log(ex);
                }

                cancellationTokenSource.Dispose();
            }

            //if (workerThread != null && workerThread.IsAlive)
            //    workerThread.Abort();
        }

        private async Task DoWorkAsync(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                try
                {
                    Globals.Log("EXX - Servicio Conector FE: Iniciado");
                    await Task.Run(() => Main.IniciarApp(), cancellationToken);
                    Globals.Log("EXX - Servicio Conector FE: Finalizado");
                }
                catch (Exception ex)
                {
                    Globals.Log("Error en ejecución del servicio: " + ex.Message);
                    Globals.Log(ex);
                }
                finally
                {
                    GC.Collect();
                }

                try
                {
                    await Task.Delay(TimeSpan.FromSeconds(10), cancellationToken);
                }
                catch (TaskCanceledException)
                {
                    break;
                }
            }

            Globals.Log("Ciclo de trabajo finalizado.");
        }

        //private void DoWork()
        //{
        //    while (true)
        //    {
        //        Main oApp;
        //        try
        //        {
        //            oApp = new Main();
        //        }
        //        catch (Exception ex)
        //        {
        //            throw ex;
        //        }
        //        finally
        //        {
        //            GC.Collect();
        //        }
        //        Thread.Sleep(10000);
        //    }
        //}
    }
}
