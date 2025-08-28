using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ServicioConectorFE.Functionality;
using ServicioConectorFE.Framework;

namespace ServicioConectorFE.Core
{
    public class Main
    {
        public static void IniciarApp()
        {
            try
            {
                Globals.GetConfig();
                for (int i = 0; i < Globals.ListConfig.Count; i++)
                {
                    try
                    {
                        SAPDI.Conectar(Globals.ListConfig[i]);
                        if (!Setup.ValidaVersion(Globals.AddOnName, Globals.AddOnVersion))
                            Structure.CrearMetaData();

                        Globals.ObtenerConfiguracionInicial();
                        Documento.EmitirDocumento();
                        Estado.Consultar();
                        Documento.EmitirBaja();
                    }
                    catch (Exception ex)
                    {
                        Globals.Log(ex);
                    }
                    finally
                    {
                        if (Globals.oCompany.Connected)
                            Globals.oCompany.Disconnect();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
