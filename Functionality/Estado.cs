using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using ServicioConectorFE.Core;
using ServicioConectorFE.Entities.Estela;
using ServicioConectorFE.Entities.SAP;
using ServicioConectorFE.Framework;
using ServicioConectorFE.IntegradorEstelaAdjunto;
using ServicioConectorFE.OSE;

namespace ServicioConectorFE.Functionality
{
    public class Estado
    {
        public static void Consultar()
        {
            try
            {
                foreach (string tipoDoc in Globals.oDocuments)
                {
                    List<Comprobante> listaComp = new List<Comprobante>();
                    listaComp = SAPDI.ListarComprobantexEmitir(DateTime.Now.AddDays(-7), DateTime.Now, tipoDoc, "01");

                    switch (Globals.ProveedorOSE)
                    {
                        case "1":
                            Bizlinks.ConsultarEstadoCPE(tipoDoc, listaComp);
                            break;
                        case "2":
                            Paperless.ConsultarEstadoCPE(tipoDoc, listaComp);
                            break;
                        case "3":
                            Estela.ConsultarEstadoCPE(tipoDoc, listaComp);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {

                throw;
            }
        }
    }
}
