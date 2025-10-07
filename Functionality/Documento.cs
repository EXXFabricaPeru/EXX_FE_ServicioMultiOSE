using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Claims;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;
using ServicioConectorFE.Core;
using ServicioConectorFE.Entities.SAP;
using ServicioConectorFE.Framework;
using ServicioConectorFE.IntegradorEstelaBaja;

namespace ServicioConectorFE.Functionality
{
    public class Documento
    {
        public static void EmitirDocumento()
        {
            try
            {
                foreach (string tipoDoc in Globals.oDocuments)
                {
                    List<Comprobante> listaComp = new List<Comprobante>();
                    listaComp = SAPDI.ListarComprobantexEmitir(DateTime.Now.AddDays(-7), DateTime.Now, tipoDoc, "00");

                    foreach (Comprobante document in listaComp.OrderBy(x => x.DocEntry))
                    {
                        bool rptaRegistrar = false;
                        string mensajeError = "";
                        String EstadoDocumento = "";
                        string Folio = "";

                        try
                        {
                            if (string.IsNullOrEmpty(document.Serie) || string.IsNullOrEmpty(document.Correlativo) || document.Correlativo == "0")
                            {
                                bool folio = SAPDI.ActualizarFolioSAP(document);
                                if (!folio) continue;
                            }

                            switch (Globals.ProveedorOSE)
                            {
                                case "1":
                                    switch (tipoDoc)
                                    {
                                        case "01":
                                        case "03":
                                        case "07":
                                        case "08":
                                            rptaRegistrar = OSE.Bizlinks.RegistrarDocumento(document.NumeroSAP, document.ObjectId, document.CodTipoDoc, document.Tabla, document.DocEntry, ref EstadoDocumento, ref Folio, ref mensajeError);
                                            break;
                                        case "09":
                                            rptaRegistrar = OSE.Bizlinks.RegistrarGuia(document.NumeroSAP, document.ObjectId, document.CodTipoDoc, document.Tabla, document.DocEntry, ref EstadoDocumento, ref Folio, ref mensajeError);
                                            break;
                                        case "20":
                                            rptaRegistrar = OSE.Bizlinks.RegistrarRetencion(document.NumeroSAP, document.ObjectId, document.CodTipoDoc, document.Tabla, document.DocEntry, ref EstadoDocumento, ref Folio, ref mensajeError);
                                            break;
                                    }
                                    break;
                                case "2":
                                    switch (tipoDoc)
                                    {
                                        case "01":
                                        case "03":
                                        case "07":
                                        case "08":
                                            rptaRegistrar = OSE.Paperless.RegistrarDocumento(document.NumeroSAP, document.ObjectId, document.CodTipoDoc, document.Tabla, document.DocEntry, ref EstadoDocumento, ref Folio, ref mensajeError);
                                            break;
                                    }
                                    break;
                                case "3":
                                    switch (tipoDoc)
                                    {
                                        case "01":
                                        case "03":
                                        case "07":
                                        case "08":
                                            rptaRegistrar = OSE.Estela.RegistrarDocumento(document.NumeroSAP, document.ObjectId, document.CodTipoDoc, document.Tabla, document.DocEntry, ref EstadoDocumento, ref Folio, ref mensajeError);
                                            break;
                                        case "09":
                                            rptaRegistrar = OSE.Estela.RegistrarGuia(document.NumeroSAP, document.ObjectId, document.CodTipoDoc, document.Tabla, document.DocEntry, ref EstadoDocumento, ref Folio, ref mensajeError);
                                            break;
                                        case "20":
                                            rptaRegistrar = OSE.Estela.RegistrarRetencion(document.NumeroSAP, document.ObjectId, document.CodTipoDoc, document.Tabla, document.DocEntry, ref EstadoDocumento, ref Folio, ref mensajeError);
                                            break;
                                    }
                                    break;
                                case "4":
                                    switch (tipoDoc)
                                    {
                                        case "01":
                                        case "03":
                                        case "07":
                                        case "08":
                                            rptaRegistrar = OSE.Nubefact.RegistrarDocumento(document.NumeroSAP, document.ObjectId, document.CodTipoDoc, document.Tabla, document.DocEntry, ref EstadoDocumento, ref Folio, ref mensajeError);
                                            break;
                                        case "09":
                                            rptaRegistrar = OSE.Nubefact.RegistrarGuia(document.NumeroSAP, document.ObjectId, document.CodTipoDoc, document.Tabla, document.DocEntry, ref EstadoDocumento, ref Folio, ref mensajeError);
                                            break;
                                        case "20":
                                            rptaRegistrar = OSE.Nubefact.RegistrarRetencion(document.NumeroSAP, document.ObjectId, document.CodTipoDoc, document.Tabla, document.DocEntry, ref EstadoDocumento, ref Folio, ref mensajeError);
                                            break;
                                    }
                                    break;

                            }

                            if (rptaRegistrar)
                            {
                                document.Nota = mensajeError;

                                switch (EstadoDocumento)
                                {
                                    case "01":
                                        document.Nota = "Documento ha sido enviado con éxito";
                                        SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, document.Nota, EstadoDocumento, Globals.oCompany);
                                        Globals.Log(document.TipoDocumento + "-" + document.Serie + "-" + document.Correlativo + ": " + document.Nota);
                                        break;
                                    default:
                                        Globals.Log(document.TipoDocumento + "-" + document.Serie + "-" + document.Correlativo + ": " + document.Nota);
                                        break;
                                }
                            }
                            else
                            {
                                SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, mensajeError, "00", Globals.oCompany);
                                Globals.Log(document.TipoDocumento + "-" + document.Serie + "-" + document.Correlativo + ": " + mensajeError);
                            }
                        }
                        catch (Exception ex)
                        {
                            Globals.Log(ex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

        public static void EmitirBaja()
        {
            try
            {
                foreach (string tipoDoc in Globals.oDocuments)
                {
                    bool rptaRegistrar = false;
                    String EstadoDocumento = "";
                    string mensaje = "";

                    List<Baja> listaComp = new List<Baja>();
                    listaComp = SAPDI.ListarComprobanteBaja(DateTime.Now.AddDays(-7), DateTime.Now, tipoDoc);

                    foreach (Baja document in listaComp)
                    {
                        try
                        {
                            switch (Globals.ProveedorOSE)
                            {
                                case "1":
                                    break;
                                case "2":
                                    break;
                                case "3":
                                    rptaRegistrar = OSE.Estela.RegistrarBaja(document, ref EstadoDocumento, ref mensaje);
                                    break;
                                case "4":

                                    break;
                            }

                            SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, mensaje, EstadoDocumento, Globals.oCompany);

                            if (document.ObjTypeAux != 0 && document.DocEntryAux != 0)
                                SAPDI.ActualizarEstadoFE(document.ObjTypeAux, document.DocEntryAux, mensaje, EstadoDocumento, Globals.oCompany);

                            Globals.Log(document.tipoComprobante + "-" + document.serie + "-" + document.correlativo + ": " + mensaje);

                            //if (rptaRegistrar)
                            //{
                            //    SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, mensaje, EstadoDocumento, Globals.oCompany);
                            //    SAPDI.ActualizarEstadoFE(document.ObjTypeAux, document.DocEntryAux, mensaje, EstadoDocumento, Globals.oCompany);
                            //}
                            //else
                            //{
                            //    SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, mensaje, document.Estado, Globals.oCompany);
                            //}
                        }
                        catch (Exception ex)
                        {
                            Globals.Log(ex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Globals.Log(ex);
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}