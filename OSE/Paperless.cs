using Microsoft.Extensions.Logging.Abstractions;
using Renci.SshNet;
using ServicioConectorFE.Core;
using ServicioConectorFE.Entities.Estela;
using ServicioConectorFE.Entities.SAP;
using ServicioConectorFE.Framework;
using ServicioConectorFE.IntegradorEstelaBaja;
using ServicioConectorFE.IntegradorPaperless;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ServicioConectorFE.OSE
{
    public class Paperless
    {
        public static bool RegistrarRetencion(string docmun, int objectype, string tipodocumento, string Tabla, int docentry, ref string EstadoDocumento, ref string Folio, ref string mensaje_error)
        {
            bool result = false;

            Entities.Estela.Retencion oDocumentoRetFE = new Entities.Estela.Retencion();

            try
            {
                DateTime date1 = DateTime.Now;
                SAPDI.ObtenerDatosCRE(docentry, ref oDocumentoRetFE);
                String Cadena = string.Empty;

                Globals.ValidaCamposCRE(oDocumentoRetFE);

                try
                {
                    //Encabezado
                    Cadena += "A;CODI_EMPR;;" + Globals.IdEmpresa + Environment.NewLine;
                    Cadena += "A;TipoDTE;;" + oDocumentoRetFE.tipoDocumento + Environment.NewLine;
                    Cadena += "A;Serie;;" + oDocumentoRetFE.serieNumeroRetencion.Split('-').ToList()[0] + Environment.NewLine;
                    Cadena += "A;Correlativo;;" + oDocumentoRetFE.serieNumeroRetencion.Split('-').ToList()[1] + Environment.NewLine;
                    Cadena += "A;FchEmis;;" + oDocumentoRetFE.fechaEmision + Environment.NewLine;
                    //Cadena += "A;HoraEmision;;00:00:00" + Environment.NewLine;
                    //Cadena += "A;TipoMoneda;;" + oDocumentoRetFE.tipoMonedaTotalPagado + Environment.NewLine;

                    //Datos del emisor
                    Cadena += "A;RUTEmis;;" + oDocumentoRetFE.numeroDocumentoEmisor + Environment.NewLine;
                    Cadena += "A;NomComer;;" + oDocumentoRetFE.razonSocialEmisor + Environment.NewLine;
                    Cadena += "A;RznSocEmis;;" + oDocumentoRetFE.razonSocialEmisor + Environment.NewLine;
                    //Cadena += "A;TipoRucEmis;;" + oDocumentoRetFE.tipoDocumentoEmisor + Environment.NewLine;
                    Cadena += "A;ComuEmis;;" + oDocumentoRetFE.ubigeoEmisor + Environment.NewLine;
                    Cadena += "A;DirEmis;;" + oDocumentoRetFE.direccionEmisor.Replace('\r', ' ').Replace(Environment.NewLine, "") + Environment.NewLine;
                    Cadena += "A;ProviEmis;;" + oDocumentoRetFE.provinciaEmisor + Environment.NewLine;
                    Cadena += "A;DeparEmis;;" + oDocumentoRetFE.departamentoEmisor + Environment.NewLine;
                    Cadena += "A;DistriEmis;;" + oDocumentoRetFE.distritoEmisor + Environment.NewLine;
                    Cadena += "A;PaisEmis;;" + oDocumentoRetFE.codigoPaisEmisor + Environment.NewLine;

                    //Datos del receptor
                    Cadena += "A;TipoRutReceptor;;" + oDocumentoRetFE.tipoDocumentoProveedor + Environment.NewLine;
                    Cadena += "A;RUTRecep;;" + oDocumentoRetFE.numeroDocumentoProveedor + Environment.NewLine;
                    Cadena += "A;RznSocRecep;;" + oDocumentoRetFE.razonSocialProveedor + Environment.NewLine;
                    Cadena += "A;DirRecepUbiGeo;;" + oDocumentoRetFE.ubigeoProveedor + Environment.NewLine;
                    Cadena += "A;DirRecep;;" + oDocumentoRetFE.direccionProveedor.Replace('\r', ' ').Replace(Environment.NewLine, "") + Environment.NewLine;
                    Cadena += "A;DirRecepProvincia;;" + oDocumentoRetFE.provinciaProveedor + Environment.NewLine;
                    Cadena += "A;DirRecepDepartamento;;" + oDocumentoRetFE.departamentoProveedor + Environment.NewLine;
                    Cadena += "A;DirRecepDistrito;;" + oDocumentoRetFE.distritoProveedor + Environment.NewLine;
                    Cadena += "A;DirRecepCodPais;;" + oDocumentoRetFE.codigoPaisProveedor + Environment.NewLine;

                    //Retención
                    Cadena += "A;CodRetencion;;" + oDocumentoRetFE.regimenRetencion + Environment.NewLine;
                    Cadena += "A;MntImpRetencion;;" + oDocumentoRetFE.tasaRetencion + Environment.NewLine;
                    Cadena += "A;ObsRetencion;;" + oDocumentoRetFE.observaciones + Environment.NewLine;
                    Cadena += "A;MntRetencion;;" + oDocumentoRetFE.importeTotalRetenido + Environment.NewLine;
                    Cadena += "A;MntTotalMenosRetencion;;" + oDocumentoRetFE.importeTotalPagado + Environment.NewLine;

                    for (int i = 0; i < oDocumentoRetFE.RetencionItems.Count; i++)
                    {
                        //Referencias
                        Cadena += "D;NroLinRef;" + (i + 1) + ";" + (i + 1) + Environment.NewLine;
                        Cadena += "D;TpoDocRef;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].tipoDocumentoRelacionado + Environment.NewLine;
                        Cadena += "D;SerieRef;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].numeroDocumentoRelacionado.Split('-').ToList()[0] + Environment.NewLine;
                        Cadena += "D;FolioRef;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].numeroDocumentoRelacionado.Split('-').ToList()[1] + Environment.NewLine;

                        //Documentos de Referencia de Comprobantes(Retención y Percepción)
                        Cadena += "D;FechEmisDocRef;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].fechaEmisionDocumentoRelacionado + Environment.NewLine;
                        Cadena += "D;MntTotalDocRef;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].importeTotalDocumentoRelacionado + Environment.NewLine;
                        Cadena += "D;MonedaDocRef;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].tipoMonedaDocumentoRelacionado + Environment.NewLine;

                        //Datos de Pago/Cobro
                        Cadena += "D;FechOperacion;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].fechaRetencion + Environment.NewLine;
                        Cadena += "D;NroOperacion;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].numeroPago + Environment.NewLine;
                        Cadena += "D;ImporteOperacion;" + (i + 1) + ";" + Globals.SumarCadenas(oDocumentoRetFE.RetencionItems[i].importeTotalPagarNeto, oDocumentoRetFE.RetencionItems[i].importeRetenido) + Environment.NewLine;
                        Cadena += "D;MonedaOperacion;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].monedaPago + Environment.NewLine;

                        Cadena += "D;ImporteMovimiento;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].importeRetenido + Environment.NewLine;
                        Cadena += "D;MonedaMovimiento;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].monedaImporteRetenido + Environment.NewLine;
                        Cadena += "D;FechaMovimiento;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].fechaRetencion + Environment.NewLine;
                        Cadena += "D;TotalMovimiento;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].importeTotalPagarNeto + Environment.NewLine;
                        Cadena += "D;Moneda;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].monedaPago + Environment.NewLine;

                        if (Convert.ToDouble(oDocumentoRetFE.RetencionItems[i].factorTipoCambioMoneda) > 1)
                        {
                            //Tipo de Cambio
                            Cadena += "D;MonedaReferencia;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].monedaReferenciaTipoCambio + Environment.NewLine;
                            Cadena += "D;MonedaObjetivo;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].monedaPago + Environment.NewLine;
                            Cadena += "D;TipoCambio;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].factorTipoCambioMoneda + Environment.NewLine;
                            Cadena += "D;FechTipoCambio;" + (i + 1) + ";" + oDocumentoRetFE.RetencionItems[i].fechaCambio + Environment.NewLine;
                        }
                    }

                    //Seccion de correos M

                    System.IO.File.WriteAllText(Globals.PathArchivos + oDocumentoRetFE.numeroDocumentoEmisor + "_" + oDocumentoRetFE.tipoDocumento + "_" + oDocumentoRetFE.serieNumeroRetencion.Replace('-', '_') + ".TXT", Cadena);

                    Folio = oDocumentoRetFE.serieNumeroRetencion;
                    //SAPDI.ActualizarEstadoFE(objectype, docentry, "Documento enviado a OSE", "01", Globals.oCompany);
                }
                catch (System.Exception ex)
                {
                }

                EstadoDocumento = "01";
                GC.Collect();
                result = true;
            }
            catch (Exception ex)
            {
                SAPDI.ActualizarEstadoFE(objectype, docentry, ex.Message, "", Globals.oCompany);
                return false;
            }
            return result;
        }

        public static bool RegistrarDocumento(string docmun, int objectype, string tipodocumento, string Tabla, int docentry, ref string EstadoDocumento, ref string Folio, ref string mensaje_error)
        {
            bool result = false;
            Entities.Paperless.Documento oBeFEDOC = new Entities.Paperless.Documento();
            System.Data.DataTable dtDetalles = new System.Data.DataTable();

            try
            {
                DateTime date1 = DateTime.Now;

                oBeFEDOC.DocEntry = docentry.ToString();
                SAPDI.ObtenerDatosCPEPaperless(ref oBeFEDOC, Tabla);
                oBeFEDOC.Detalles = new List<Entities.Paperless.DocLine>();
                SAPDI.ObtenerDetalleDocumento(ref dtDetalles, Tabla, docentry);

                String Cadena = string.Empty;

                ValidaCamposCPE(oBeFEDOC);

                //Montos
                Double MntGrabado = 0;
                Double MntInafecto = 0;
                Double MntExonerado = 0;
                Double MntTransGratuita = 0;
                Double MntImpTransGrat = 0;
                Double MntIGV = 0;
                Double MntDescuentos = 0;
                Double MntTotal = Convert.ToDouble(oBeFEDOC.importeTotal);
                Double MntDescuentoGlobal = Convert.ToDouble(oBeFEDOC.valorDescuentoGlobal);

                #region Calcula totales
                //Asigna código de moneda
                switch (oBeFEDOC.tipoMoneda)
                {
                    case "PEN":
                    case "SOL":
                        oBeFEDOC.tipoMoneda = "PEN";
                        break;
                    case "US$":
                    case "USD":
                        oBeFEDOC.tipoMoneda = "USD";
                        break;
                    case "EUR":
                        oBeFEDOC.tipoMoneda = "EUR";
                        break;
                }
                if (oBeFEDOC.tipoMoneda == "PEN")
                    oBeFEDOC.TipoCambio = "0";
                Globals.TipoCambio = Convert.ToDouble(oBeFEDOC.TipoCambio);


                oBeFEDOC.Detalles = new List<Entities.Paperless.DocLine>();
                int RowsinMatrix = 1;

                if (oBeFEDOC.DocType == "I") //Articulos
                {
                    foreach (System.Data.DataRow detalle in dtDetalles.Rows)
                    {
                        Entities.Paperless.DocLine oDetalle = new Entities.Paperless.DocLine();
                        oDetalle.posicion = RowsinMatrix.ToString();
                        oDetalle.codigo = detalle.Field<string>("codigoArt").ToString();
                        oDetalle.descripcion = detalle.Field<string>("descripcion").ToString();
                        oDetalle.cantidadUnidades = Globals.RetirarMonedayDecimal(detalle.Field<string>("cantidadUnidades"));
                        oDetalle.unidadMedidaSAP = detalle.Field<string>("unidadMedidaSAP").ToString();
                        oDetalle.unidadMedida = detalle.Field<string>("unidadMedida").ToString();
                        oDetalle.valorUnitario = Globals.RetirarMonedayDecimal(detalle.Field<string>("valorUnitario").ToString(), 6);
                        oDetalle.IgvMonto = Globals.RetirarMonedayDecimal(detalle.Field<string>("montoIGV").ToString());
                        oDetalle.precioUnitario = Globals.RetirarMonedayDecimal(detalle.Field<string>("precioUnitario").ToString());
                        oDetalle.valorItem = Globals.RetirarMonedayDecimal(detalle.Field<string>("valorItem").ToString());
                        oDetalle.CodigoImpuesto = detalle.Field<string>("CodigoImpuesto").ToString();
                        oDetalle.tasaIGV = Convert.ToDouble(detalle.Field<string>("tasaIGV"));
                        oDetalle.Rate = Globals.RetirarMonedayDecimal(detalle.Field<string>("Rate").ToString());
                        oDetalle.IgvTipo = detalle.Field<string>("U_SMC_TIPOAFECT_FE").ToString();
                        oDetalle.descuentoPct = Globals.RetirarMonedayDecimal(detalle.Field<string>("DiscPrcnt"));
                        oDetalle.OperacionAfecta = oDetalle.CodigoImpuesto == "IGV" ? "A" : oDetalle.CodigoImpuesto == "INA" ? "I" : oDetalle.CodigoImpuesto == "EXO" ? "E" : ""; // detalle.Field<string>("U_BPP_OPER").ToString();
                        oDetalle.TaxOnly = detalle.Field<string>("TaxOnly").ToString();
                        oDetalle.Operacion = detalle.Field<string>("Operacion").ToString();
                        oDetalle.CondicionCarga = detalle.Field<string>("CondicionCarga").ToString();
                        oDetalle.Tamano = detalle.Field<string>("Tamaño").ToString();

                        Double precioUnitario = 0;
                        Double Rate = Convert.ToDouble(oDetalle.Rate);
                        Double valorUnitario = Convert.ToDouble(oDetalle.valorUnitario);
                        if (oDetalle.descuentoPct == "0.00")
                            precioUnitario = Convert.ToDouble(oDetalle.precioUnitario);
                        else
                            precioUnitario = Math.Round((valorUnitario * (1 + oDetalle.tasaIGV / 100)), 2, MidpointRounding.AwayFromZero);

                        if (Rate > 0 && oBeFEDOC.tipoMoneda == "PEN")//Precio item en Moneda extranjera y Documento en Soles
                        {
                            valorUnitario = Math.Round(Convert.ToDouble(valorUnitario) * Rate, 6, MidpointRounding.AwayFromZero);
                            precioUnitario = Math.Round(Convert.ToDouble(precioUnitario) * Rate, 2, MidpointRounding.AwayFromZero);
                        }
                        else if (Rate == 0 && oBeFEDOC.tipoMoneda != "PEN")//Precio de item en Soles y Documento en Moneda extranjera
                        {
                            valorUnitario = Math.Round(Convert.ToDouble(valorUnitario) / Globals.TipoCambio, 6, MidpointRounding.AwayFromZero);
                            precioUnitario = Math.Round(Convert.ToDouble(precioUnitario) / Globals.TipoCambio, 2, MidpointRounding.AwayFromZero);
                        }
                        oDetalle.valorUnitario = Globals.RetirarMonedayDecimal(valorUnitario.ToString(), 6);
                        oDetalle.precioUnitario = Globals.RetirarMonedayDecimal(precioUnitario.ToString());

                        if (oBeFEDOC.TransGratuita)
                        {
                            oDetalle.ValorCodigo = "";
                            oDetalle.CodigoTributo = "9996";
                            oDetalle.CodigoTributoInt = "FRE";
                            oDetalle.NombreTributo = "GRA";
                            MntTransGratuita += (precioUnitario * Convert.ToDouble(oDetalle.cantidadUnidades));
                            MntImpTransGrat += Convert.ToDouble(oDetalle.IgvMonto);
                            oDetalle.valorItemGratutito = oDetalle.valorItem;
                            oDetalle.IgvMontoGratuito = oDetalle.IgvMonto;
                            oDetalle.valorUnitario = "0.00";
                            oDetalle.valorItem = "0.00";
                            oDetalle.IgvMonto = "0.00";
                            if (string.IsNullOrEmpty(oDetalle.IgvTipo)) throw new System.Exception("Para transferencia gratuita debe seleccionar el campo FE TIPO OPERACION en los detalles del documento");
                        }
                        else
                        {
                            oDetalle.valorItemGratutito = "0.00";
                            oDetalle.IgvMontoGratuito = "0.00";
                            if (oBeFEDOC.documentoExportacion == "Y")
                            {
                                oDetalle.ValorCodigo = "";
                                oDetalle.CodigoTributo = "9995";
                                oDetalle.CodigoTributoInt = "FRE";
                                oDetalle.NombreTributo = "EXP";
                                oDetalle.IgvTipo = "40";
                                MntExonerado += Convert.ToDouble(oDetalle.valorItem);
                            }
                            else
                            {
                                if (oDetalle.OperacionAfecta == "E")
                                {
                                    oDetalle.ValorCodigo = "E";
                                    oDetalle.CodigoTributo = "9997";
                                    oDetalle.CodigoTributoInt = "VAT";
                                    oDetalle.NombreTributo = "EXO";
                                    oDetalle.IgvTipo = "20";
                                    MntExonerado += Convert.ToDouble(oDetalle.valorItem);
                                }
                                else if (oDetalle.OperacionAfecta == "I")
                                {
                                    oDetalle.ValorCodigo = "O";
                                    oDetalle.CodigoTributo = "9998";
                                    oDetalle.CodigoTributoInt = "FRE";
                                    oDetalle.NombreTributo = "INA";
                                    oDetalle.IgvTipo = "30";
                                    MntInafecto += Convert.ToDouble(oDetalle.valorItem);
                                }
                                else if (oDetalle.OperacionAfecta == "A")
                                {
                                    oDetalle.ValorCodigo = "S";
                                    oDetalle.CodigoTributo = "1000";
                                    oDetalle.CodigoTributoInt = "VAT";
                                    oDetalle.NombreTributo = "IGV";
                                    oDetalle.IgvTipo = "10";
                                    MntGrabado += Convert.ToDouble(oDetalle.valorItem);
                                    MntIGV += Convert.ToDouble(oDetalle.IgvMonto);
                                }
                            }
                        }

                        Double CalcDescuento = Math.Round(
                                               (
                                               Convert.ToDouble(oDetalle.cantidadUnidades)
                                               *
                                               Convert.ToDouble(oDetalle.valorUnitario)
                                               )
                                               *
                                               (
                                               Convert.ToDouble(Globals.RetirarMonedayDecimal(oDetalle.descuentoPct))
                                               *
                                               0.01
                                               ), 2, MidpointRounding.AwayFromZero);
                        oDetalle.descuento = Globals.RetirarMonedayDecimal(CalcDescuento.ToString());
                        if (Convert.ToDouble(oDetalle.descuento) < 0)
                            oDetalle.descuento = (Convert.ToDouble(oDetalle.descuento) * -1).ToString();
                        MntDescuentos += Convert.ToDouble(oDetalle.descuento);

                        oBeFEDOC.Detalles.Add(oDetalle);
                        RowsinMatrix++;
                    }
                    //MntIGV = MntTotal - MntGrabado - MntInafecto - MntExonerado;
                }
                else //Servicios
                {
                    foreach (System.Data.DataRow detalle in dtDetalles.Rows)
                    {
                        Entities.Paperless.DocLine oDetalle = new Entities.Paperless.DocLine();
                        oDetalle.posicion = RowsinMatrix.ToString();
                        oDetalle.codigo = detalle.Field<string>("descripcionServ").ToString();
                        oDetalle.descripcion = detalle.Field<string>("descripcion").ToString();
                        oDetalle.unidadMedidaSAP = detalle.Field<string>("unidadMedidaSAP").ToString();
                        oDetalle.unidadMedida = "NIU";
                        oDetalle.cantidadUnidades = "1.00";
                        oDetalle.valorUnitario = Globals.RetirarMonedayDecimal(detalle.Field<string>("valorUnitario").ToString(), 6);
                        oDetalle.IgvMonto = Globals.RetirarMonedayDecimal(detalle.Field<string>("montoIGV").ToString());
                        oDetalle.precioUnitario = Globals.RetirarMonedayDecimal(detalle.Field<string>("precioUnitario").ToString());
                        oDetalle.valorItem = Globals.RetirarMonedayDecimal(detalle.Field<string>("valorItem").ToString());
                        oDetalle.CodigoImpuesto = detalle.Field<string>("CodigoImpuesto").ToString();
                        oDetalle.tasaIGV = Convert.ToDouble(detalle.Field<string>("tasaIGV"));
                        oDetalle.Rate = Globals.RetirarMonedayDecimal(detalle.Field<string>("Rate").ToString());
                        oDetalle.IgvTipo = detalle.Field<string>("U_SMC_TIPOAFECT_FE").ToString();
                        oDetalle.descuentoPct = Globals.RetirarMonedayDecimal(detalle.Field<string>("DiscPrcnt"));
                        //oDetalle.OperacionAfecta = detalle.Field<string>("U_BPP_OPER").ToString();
                        oDetalle.OperacionAfecta = oDetalle.CodigoImpuesto == "IGV" ? "A" : oDetalle.CodigoImpuesto == "INA" ? "I" : oDetalle.CodigoImpuesto == "EXO" ? "E" : ""; // detalle.Field<string>("U_BPP_OPER").ToString();
                        oDetalle.TaxOnly = detalle.Field<string>("TaxOnly").ToString();
                        oDetalle.Operacion = detalle.Field<string>("Operacion").ToString();
                        oDetalle.CondicionCarga = detalle.Field<string>("CondicionCarga").ToString();
                        oDetalle.Tamano = detalle.Field<string>("Tamaño").ToString();

                        Double precioUnitario = 0;
                        Double Rate = Convert.ToDouble(oDetalle.Rate);
                        Double valorUnitario = Convert.ToDouble(oDetalle.valorUnitario);

                        if (oDetalle.descuentoPct == "0.00")
                            precioUnitario = Convert.ToDouble(oDetalle.precioUnitario);
                        else
                            precioUnitario = Math.Round((valorUnitario * (1 + oDetalle.tasaIGV / 100)), 2, MidpointRounding.AwayFromZero);

                        if (Rate > 0 && oBeFEDOC.tipoMoneda == "PEN")//Precio item en Moneda extranjera y Documento en Soles
                        {
                            valorUnitario = Math.Round(Convert.ToDouble(valorUnitario) * Rate, 6, MidpointRounding.AwayFromZero);
                            precioUnitario = Math.Round(Convert.ToDouble(precioUnitario) * Rate, 2, MidpointRounding.AwayFromZero);
                        }
                        else if (Rate == 0 && oBeFEDOC.tipoMoneda != "PEN")//Precio de item en Soles y Documento en Moneda extranjera
                        {
                            valorUnitario = Math.Round(Convert.ToDouble(valorUnitario) / Globals.TipoCambio, 6, MidpointRounding.AwayFromZero);
                            precioUnitario = Math.Round(Convert.ToDouble(precioUnitario) / Globals.TipoCambio, 2, MidpointRounding.AwayFromZero);
                        }
                        oDetalle.valorUnitario = Globals.RetirarMonedayDecimal(valorUnitario.ToString());
                        oDetalle.precioUnitario = Globals.RetirarMonedayDecimal(precioUnitario.ToString());

                        if (oBeFEDOC.TransGratuita)
                        {
                            oDetalle.ValorCodigo = "";
                            oDetalle.CodigoTributo = "9996";
                            oDetalle.CodigoTributoInt = "FRE";
                            oDetalle.NombreTributo = "GRA";
                            MntTransGratuita += (precioUnitario * Convert.ToDouble(oDetalle.cantidadUnidades));
                            MntImpTransGrat += Convert.ToDouble(oDetalle.IgvMonto);
                            oDetalle.valorItemGratutito = oDetalle.valorItem;
                            oDetalle.IgvMontoGratuito = oDetalle.IgvMonto;
                            oDetalle.valorUnitario = "0.00";
                            oDetalle.valorItem = "0.00";
                            oDetalle.IgvMonto = "0.00";
                            if (string.IsNullOrEmpty(oDetalle.IgvTipo)) throw new System.Exception("Para transferencia gratuita debe seleccionar el campo FE TIPO OPERACION en los detalles del documento");
                        }
                        else
                        {
                            oDetalle.valorItemGratutito = "0.00";
                            oDetalle.IgvMontoGratuito = "0.00";
                            if (oBeFEDOC.documentoExportacion == "Y")
                            {
                                oDetalle.ValorCodigo = "";
                                oDetalle.CodigoTributo = "9995";
                                oDetalle.CodigoTributoInt = "FRE";
                                oDetalle.NombreTributo = "EXP";
                                oDetalle.IgvTipo = "40";
                                MntExonerado += Convert.ToDouble(oDetalle.valorItem);
                            }
                            else
                            {
                                if (oDetalle.OperacionAfecta == "E")
                                {
                                    oDetalle.ValorCodigo = "E";
                                    oDetalle.CodigoTributo = "9997";
                                    oDetalle.CodigoTributoInt = "VAT";
                                    oDetalle.NombreTributo = "EXO";
                                    oDetalle.IgvTipo = "20";
                                    MntExonerado += Convert.ToDouble(oDetalle.valorItem);
                                }
                                else if (oDetalle.OperacionAfecta == "I")
                                {
                                    oDetalle.ValorCodigo = "O";
                                    oDetalle.CodigoTributo = "9998";
                                    oDetalle.CodigoTributoInt = "FRE";
                                    oDetalle.NombreTributo = "INA";
                                    oDetalle.IgvTipo = "30";
                                    MntInafecto += Convert.ToDouble(oDetalle.valorItem);
                                }
                                else if (oDetalle.OperacionAfecta == "A")
                                {
                                    oDetalle.ValorCodigo = "S";
                                    oDetalle.CodigoTributo = "1000";
                                    oDetalle.CodigoTributoInt = "VAT";
                                    oDetalle.NombreTributo = "IGV";
                                    oDetalle.IgvTipo = "10";
                                    MntGrabado += Convert.ToDouble(oDetalle.valorItem);
                                    MntIGV += Convert.ToDouble(oDetalle.IgvMonto);
                                }
                            }
                        }

                        Double CalcDescuento = Math.Round(
                                               (
                                               Convert.ToDouble(oDetalle.cantidadUnidades)
                                               *
                                               Convert.ToDouble(oDetalle.valorUnitario)
                                               )
                                               *
                                               (
                                               Convert.ToDouble(Globals.RetirarMonedayDecimal(oDetalle.descuentoPct))
                                               *
                                               0.01
                                               ), 2, MidpointRounding.AwayFromZero);
                        oDetalle.descuento = Globals.RetirarMonedayDecimal(CalcDescuento.ToString());
                        if (Convert.ToDouble(oDetalle.descuento) < 0)
                            oDetalle.descuento = (Convert.ToDouble(oDetalle.descuento) * -1).ToString();
                        MntDescuentos += Convert.ToDouble(oDetalle.descuento);

                        oBeFEDOC.Detalles.Add(oDetalle);
                        RowsinMatrix++;
                    }
                    //MntIGV = MntTotal - MntGrabado - MntInafecto - MntExonerado;
                }
                #endregion

                #region MontosTotales
                oBeFEDOC.tasaIgv = "0.18";// Globals.ObtenerTasaIGV();

                if (Tabla == "ODPI")
                {
                    oBeFEDOC.valorDescuentoGlobal = "0.00";
                    MntGrabado = Convert.ToDouble(oBeFEDOC.valorTotalAnticipos);
                    oBeFEDOC.valorTotalAnticipos = "0.00";
                    oBeFEDOC.valorIgv = Globals.RetirarMonedayDecimal((MntGrabado * Convert.ToDouble(oBeFEDOC.tasaIgv)).ToString());
                }
                else
                {
                    oBeFEDOC.valorDescuentoGlobal = Globals.RetirarMonedayDecimal((MntDescuentoGlobal).ToString());
                    oBeFEDOC.valorIgv = Globals.RetirarMonedayDecimal((MntGrabado * Convert.ToDouble(oBeFEDOC.tasaIgv)).ToString());
                    oBeFEDOC.valorIgv = Globals.SumarCadenas(oBeFEDOC.valorIgv, "-" + oBeFEDOC.valorTotalIGVAnticipos);
                    MntGrabado = MntGrabado - (Convert.ToDouble(oBeFEDOC.valorTotalAnticipos) + MntDescuentoGlobal);

                    if (Convert.ToDouble(oBeFEDOC.valorTotalAnticipos) > 0)
                        Globals.ObtenerAnticiposPaperless(ref oBeFEDOC, Tabla, Convert.ToInt32(oBeFEDOC.DocEntry));
                }

                oBeFEDOC.valorOperacionesGravadas = Globals.RetirarMonedayDecimal(MntGrabado.ToString());
                oBeFEDOC.valorOperacionesExoneradas = Globals.RetirarMonedayDecimal(MntExonerado.ToString());
                oBeFEDOC.valorOperacionesInafectas = Globals.RetirarMonedayDecimal(MntInafecto.ToString());
                oBeFEDOC.valorOperacionesGratuitas = Globals.RetirarMonedayDecimal(MntTransGratuita.ToString());

                oBeFEDOC.otrosCargos = "0.00";

                if (MntImpTransGrat > 0)
                {
                    if (MntGrabado == 0)
                    {
                        oBeFEDOC.valorOperacionesGravadas = "0.00";
                        oBeFEDOC.valorIgv = "0.00";
                        oBeFEDOC.importeTotal = "0.00";
                    }
                    else
                    {
                        oBeFEDOC.valorIgv = Globals.RetirarMonedayDecimal(Globals.SumarCadenas(oBeFEDOC.valorIgv, "-" + oBeFEDOC.valorTotalIGVAnticipos).ToString());

                        //(Convert.ToDouble(UtilCS.oBEiFEDoc.totales.valorIgv) - MntImpTransGrat).ToString();
                        oBeFEDOC.importeTotal = Globals.RetirarMonedayDecimal(Globals.SumarCadenas(oBeFEDOC.importeTotal, (MntImpTransGrat * -1).ToString()));
                        //(Convert.ToDouble(UtilCS.oBEiFEDoc.totales.importeTotal) - MntImpTransGrat).ToString();
                    }
                }

                //Gastos Adicionales
                //Globals.ObtenerGastosAdicionales(ref oBeFEDOC, Convert.ToInt32(oBeFEDOC.DocEntry), Globals.SAPTable);
                #endregion

                #region SECCION EN: ENCABEZADO
                oBeFEDOC.valorIgv = Convert.ToDouble(oBeFEDOC.valorIgv) < 0 ? "0.00" : oBeFEDOC.valorIgv;
                Cadena = "<![CDATA[[" + Environment.NewLine;
                Cadena += string.Format("EN|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}|{19}|{20}|{21}|{22}|{23}|{24}|{25}|{26}|{27}|{28}|{29}|{30}",
                oBeFEDOC.tipoComprobante,
                (oBeFEDOC.serie + "-" + oBeFEDOC.correlativo),
                oBeFEDOC.notaTipo,
                oBeFEDOC.ListReferencia.Count > 0 ? oBeFEDOC.ListReferencia[0].serie + "-" + oBeFEDOC.ListReferencia[0].correlativo : "",
                oBeFEDOC.notaSustento,
                oBeFEDOC.fechaEmision,
                oBeFEDOC.tipoMoneda,
                oBeFEDOC.emisor.ruc,
                oBeFEDOC.emisor.tipoIdentificacion,
                oBeFEDOC.emisor.nombreComercial,
                oBeFEDOC.emisor.razonSocial,
                oBeFEDOC.emisor.ubigeo,
                oBeFEDOC.emisor.direccion,
                oBeFEDOC.emisor.departamento,
                oBeFEDOC.emisor.provincia,
                oBeFEDOC.emisor.distrito,
                oBeFEDOC.adquiriente.numeroIdentificacion,
                oBeFEDOC.adquiriente.tipoIdentificacion,
                oBeFEDOC.adquiriente.nombre,
                oBeFEDOC.adquiriente.direccionFiscal,
                //oBeFEDOC.importeTotal == "0.00" ? Utils.SumarCadenas(oBeFEDOC.importeTotal, (Convert.ToDouble(oBeFEDOC.valorIgv) * -1).ToString())
                //                                      : Utils.SumarCadenas(Utils.SumarCadenas(oBeFEDOC.importeTotal, (Convert.ToDouble(oBeFEDOC.valorIgv) * -1).ToString()), (Convert.ToDouble(oBeFEDOC.redondeo) * -1).ToString()),
                Globals.SumarCadenas(Globals.SumarCadenas(oBeFEDOC.importeTotal, (Convert.ToDouble(oBeFEDOC.valorIgv) * -1).ToString()), oBeFEDOC.valorTotalAnticipos),
                oBeFEDOC.valorIgv,
                oBeFEDOC.valorDescuentoGlobal,
                oBeFEDOC.otrosCargos,
                Globals.SumarCadenas(oBeFEDOC.importeTotal, "-" + Globals.SumarCadenas(oBeFEDOC.valorDescuentoGlobal, oBeFEDOC.otrosCargos)),
                string.Empty, //Otros conceptos tributarios o comerciales, se envia valor en blanco de acuerdo a documentación
                string.Empty, // Total de valor venta neto, se envia valor en blanco de acuerdo a documentación
                oBeFEDOC.adquiriente.numeroIdentificacion,
                oBeFEDOC.adquiriente.tipoIdentificacion,
                oBeFEDOC.adquiriente.codigoPais,
                string.Empty
                ) + Environment.NewLine;
                #endregion

                #region SECCION ENEX: ENCABEZADO EXTENSION
                Cadena += string.Format("ENEX|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}",
                "2.1",
                oBeFEDOC.tipoOperacion,
                string.IsNullOrEmpty(oBeFEDOC.ordenCompra) ? "" : oBeFEDOC.ordenCompra,
                oBeFEDOC.redondeo,
                //Utils.SumarCadenas(oBeFEDOC.valorTotalAnticipos, oBeFEDOC.valorTotalIGVAnticipos),
                Globals.SumarCadenas(oBeFEDOC.valorTotalAnticipos, oBeFEDOC.valorTotalIGVAnticipos),
                oBeFEDOC.fechaVencimiento,
                oBeFEDOC.horaEmision,
                oBeFEDOC.codigoEstablecimiento,
                Globals.SumarCadenas(oBeFEDOC.importeTotal, (Globals.SumarCadenas(oBeFEDOC.valorTotalAnticipos, oBeFEDOC.valorTotalIGVAnticipos))), //UtilCS.oBEiFEDoc.montoTotal,
                string.Empty, //Aplica para venta de vehiculos
                string.Empty, //Aplica para venta de vehiculos
                string.Empty  //Aplica para venta de vehiculos
                ) + Environment.NewLine;
                #endregion

                #region SECCION FGA: FACTURA GUIA
                //No aplica Factura guía
                //if (pVal.FormTypeEx == UtilCS.iFEFormularios.Entrega.ToString())
                //{
                //    Cadena += string.Format("FGA|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}|{15}|{16}|{17}|{18}|{19}|{20}|{21}|{22}|{23}|{24}|{25}|{26}|{27}|{28}|{29}|{30}",
                //    "");
                //}
                #endregion

                #region SECCION VITE: VENTA ITINERANTE
                //No aplica Venta itinerante
                //if (UtilCS.oBEiFEDoc.tipoComprobante == "0104" || UtilCS.oBEiFEDoc.tipoComprobante == "0201" || UtilCS.oBEiFEDoc.tipoComprobante == "0208")
                //{
                //    Cadena += string.Format("VITE|{0}|{1}|{2}|{3}|{4}|{5}|{6}",
                //    "");
                //}
                #endregion

                #region SECCION FEX: FACTURA EXPORTACION
                //No aplica Venta de exportación
                #endregion

                #region SECCION PERC: PERCEPCION
                //No aplica Percepciones
                #endregion

                #region SECCION DET: DETRACCION
                if (oBeFEDOC.tipoComprobante != "07")
                {
                    if (oBeFEDOC.TieneDetraccion == "Y")
                    {
                        Cadena += string.Format("DET|{0}|{1}|{2}|{3}|{4}|{5}|||",
                                  oBeFEDOC.codigoDetraccion,
                                  "00638013932",
                                  Convert.ToDouble(oBeFEDOC.porcentajeDetraccion) * 100,
                                  oBeFEDOC.totalDetraccionMD,
                                  oBeFEDOC.tipoMoneda,
                                  "001") + Environment.NewLine;
                    }
                }
                #endregion

                #region SECCION DN: NOTAS DEL DOCUMENTO
                int indexLeyenda = 1;
                Cadena += string.Format("DN|{0}|{1}|{2}",
                indexLeyenda,
                "1000", //Codigo Catalogo 52
                Globals.MontoEnLetras(oBeFEDOC.importeTotal, Globals.ObtenerDescripcionMoneda(oBeFEDOC.tipoMoneda))) + Environment.NewLine;
                indexLeyenda++;

                if (oBeFEDOC.valorOperacionesGratuitas != "0.00")
                {
                    Cadena += string.Format("DN|{0}|{1}|{2}",
                    indexLeyenda,
                    "1002", //Codigo Catalogo 52
                    "TRANSFERENCIA GRATUITA DE UN BIEN Y / O SERVICIO PRESTADO GRATUITAMENTE") + Environment.NewLine;
                    indexLeyenda++;
                }

                //if (oBeFEDOC.TieneDetraccion == "Y")
                //{
                //    Cadena += string.Format("DN|{0}|{1}|{2}",
                //              indexLeyenda,
                //              "2006", //Codigo Catalogo 52
                //              "Operación sujeta a detracción") + Environment.NewLine;
                //    indexLeyenda++;
                //}
                #endregion

                #region PE: PERSONALIZADOS
                if (oBeFEDOC.tipoComprobante == "01" || oBeFEDOC.tipoComprobante == "03")
                {
                    if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_NAVE))
                        Cadena += string.Format("PE|{0}|{1}|", "Nave", oBeFEDOC.U_EXT_NAVE) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_CALLNUMBER))
                        Cadena += string.Format("PE|{0}|{1}|", "CallNumb", oBeFEDOC.U_EXT_CALLNUMBER) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_DESRESOLIDA))
                    {
                        if (oBeFEDOC.tipoComprobante == "01")
                            Cadena += string.Format("PE|{0}|{1}|", "RapSolid", oBeFEDOC.U_EXT_DESRESOLIDA) + Environment.NewLine;
                        else if (oBeFEDOC.tipoComprobante == "03")
                            Cadena += string.Format("PE|{0}|{1}|", "SolicitanteServicio", oBeFEDOC.U_EXT_DESRESOLIDA) + Environment.NewLine;
                    }
                    if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_OSNRO))
                        if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_OSNRO))
                            Cadena += string.Format("PE|{0}|{1}|", "NroOs", oBeFEDOC.U_EXT_OSNRO) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_COMENTARIO))
                        Cadena += string.Format("PE|{0}|{1}|", "Comentarios", oBeFEDOC.U_EXT_COMENTARIO) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_MUELLE))
                        Cadena += string.Format("PE|{0}|{1}|", "Muelle", oBeFEDOC.U_EXT_MUELLE) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_VIAJE))
                        Cadena += string.Format("PE|{0}|{1}|", "Viaje", oBeFEDOC.U_EXT_VIAJE) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_FECARRIBO))
                        if (Convert.ToDateTime(oBeFEDOC.U_EXT_FECARRIBO).Year > 2020)
                            Cadena += string.Format("PE|{0}|{1}|", "Arribo", oBeFEDOC.U_EXT_FECARRIBO) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_LINEANAV))
                        Cadena += string.Format("PE|{0}|{1}|", "Linea", oBeFEDOC.U_EXT_LINEANAV) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_RESPSOLIDA))
                        Cadena += string.Format("PE|{0}|{1}|", "Codigo", oBeFEDOC.U_EXT_RESPSOLIDA) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_dc_fact_VBWONbr) && oBeFEDOC.U_EXT_dc_fact_VBWONbr != "N")
                        Cadena += string.Format("PE|{0}|{1}|", "Servicio", oBeFEDOC.U_EXT_dc_fact_VBWONbr) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(Globals.oOADM.Phone1))
                        Cadena += string.Format("PE|{0}|{1}|", "Telefono", Globals.oOADM.Phone1) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(Globals.oOADM.Fax))
                        Cadena += string.Format("PE|{0}|{1}|", "Fax", Globals.oOADM.Fax) + Environment.NewLine;

                    Cadena += string.Format("PE|{0}|{1}|", "FormaPago", oBeFEDOC.condicionPago) + Environment.NewLine;
                    Cadena += string.Format("PE|{0}|{1}|", "FechaVencimiento", oBeFEDOC.fechaVencimiento) + Environment.NewLine;
                    Cadena += "PE|URL|https://www.puertopaita.com|" + Environment.NewLine;
                }


                if (oBeFEDOC.tipoComprobante == "07")
                {
                    if (!string.IsNullOrEmpty(oBeFEDOC.U_EXT_COMENTARIO))
                        Cadena += string.Format("PE|{0}|{1}|", "Comentarios", oBeFEDOC.U_EXT_COMENTARIO) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(Globals.oOADM.Phone1))
                        Cadena += string.Format("PE|{0}|{1}|", "Telefono", Globals.oOADM.Phone1) + Environment.NewLine;
                    if (!string.IsNullOrEmpty(Globals.oOADM.Fax))
                        Cadena += string.Format("PE|{0}|{1}|", "Fax", Globals.oOADM.Fax) + Environment.NewLine;
                    Cadena += string.Format("PE|{0}|{1}|", "URL", "https://www.puertopaita.com") + Environment.NewLine;
                    Cadena += "PE|URL|https://www.puertopaita.com|" + Environment.NewLine;
                }
                #endregion

                #region SECCION PP: ANTICIPOS Y PREPAGOS
                if (oBeFEDOC.ListAnticipo != null)
                {
                    string DescripcionAnticipo = string.Empty;
                    //string ValorAnticipo = "0.00";

                    for (int i = 0; i < oBeFEDOC.ListAnticipo.Count; i++)
                    {
                        DescripcionAnticipo += oBeFEDOC.ListAnticipo[i].serieDocumento + "-" + oBeFEDOC.ListAnticipo[i].correlativoDocumento + " ";
                        //ValorAnticipo = Utils.SumarCadenas(oBeFEDOC.ListAnticipo[i].totalAnticipos, ValorAnticipo);
                    }

                    Cadena += string.Format("PE|{0}|{1}", "CodigoAnticipo", "FACT_ANTIC") + Environment.NewLine;
                    Cadena += string.Format("PE|{0}|{1}", "DescripcionAnticipo", "AVE/" + DescripcionAnticipo.Trim()) + Environment.NewLine;
                    Cadena += string.Format("PE|{0}|{1}", "UnidadAnticipo", "UND") + Environment.NewLine;
                    Cadena += string.Format("PE|{0}|{1}", "ValorAnticipo", "-" + oBeFEDOC.valorTotalAnticipos) + Environment.NewLine;
                }

                #endregion

                #region ITEM
                for (int i = 0; i < oBeFEDOC.Detalles.Count; i++)
                {
                    #region DE: DETALLE ITEM
                    Cadena += string.Format("DE|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}",
                    oBeFEDOC.Detalles[i].posicion,
                    oBeFEDOC.Detalles[i].precioUnitario,
                    oBeFEDOC.Detalles[i].unidadMedida,
                    oBeFEDOC.Detalles[i].cantidadUnidades,
                    oBeFEDOC.Detalles[i].valorItemGratutito == "0.00" ? oBeFEDOC.Detalles[i].valorItem : oBeFEDOC.Detalles[i].valorItemGratutito,
                    oBeFEDOC.Detalles[i].codigo,
                    oBeFEDOC.Detalles[i].valorUnitario == "0.00" ? "02" : "01",
                    oBeFEDOC.Detalles[i].valorUnitario,
                    oBeFEDOC.Detalles[i].valorItemGratutito == "0.00" ? oBeFEDOC.Detalles[i].valorItem : oBeFEDOC.Detalles[i].valorItemGratutito,
                    string.Empty, //Consultar si se manejará Lotes
                    string.Empty,//No aplica indica datos de vehiculo
                    string.Empty,
                    string.Empty) + Environment.NewLine;
                    #endregion

                    #region DEDI: DESCRIPCION DEL ITEM
                    Cadena += string.Format("DEDI|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}",
                    oBeFEDOC.Detalles[i].descripcion,
                    string.Empty, //Descripcion adicional
                    string.Empty, //Catalogo 55, llenado de acuerdo a documentación
                    string.Empty, //Catalogo 55, llenado de acuerdo a documentación
                    string.Empty, //Número de placa del vehículo ?
                    string.Empty, //Confirmar con Ventura el nombre del  campo código de articulo sunat) + Environment.NewLine; 
                    string.Empty, //No se usará código GS1
                    string.Empty  //No se usará tipo de estructura
                    ) + Environment.NewLine;

                    Cadena += string.Format("DEDI|{0}", string.IsNullOrEmpty(oBeFEDOC.Detalles[i].Operacion) ? "*" : oBeFEDOC.Detalles[i].Operacion) + Environment.NewLine;
                    Cadena += string.Format("DEDI|{0}", string.IsNullOrEmpty(oBeFEDOC.Detalles[i].CondicionCarga) ? "*" : oBeFEDOC.Detalles[i].CondicionCarga) + Environment.NewLine;
                    Cadena += string.Format("DEDI|{0}", string.IsNullOrEmpty(oBeFEDOC.Detalles[i].Tamano) ? "*" : oBeFEDOC.Detalles[i].Tamano) + Environment.NewLine;
                    Cadena += string.Format("DEDI|{0}", "*") + Environment.NewLine;
                    Cadena += string.Format("DEDI|{0}", string.IsNullOrEmpty(oBeFEDOC.Detalles[i].unidadMedidaSAP) ? "*" : oBeFEDOC.Detalles[i].unidadMedidaSAP) + Environment.NewLine;
                    #endregion

                    #region DEDR: DESCUENTOS Y RECARGOS DEL ITEM
                    if (oBeFEDOC.Detalles[i].descuento != "0.00")
                    {
                        Cadena += string.Format("DEDR|{0}|{1}|{2}|{3}|{4}",
                        "FALSE",
                        oBeFEDOC.Detalles[i].descuento,
                        "00", //Catalogo 53, otros descuentos
                        Convert.ToDouble(oBeFEDOC.Detalles[i].descuentoPct) * 0.01,
                        oBeFEDOC.Detalles[i].valorItem
                        ) + Environment.NewLine;
                    }
                    //if (UtilCS.oBEiFEDoc.Detalles[i].descuento != "0.00") //En SAP no hay recargos a nivel de item
                    //{

                    //}
                    #endregion

                    #region DEIM: IMPUESTOS DEL ITEM
                    //if (oBeFEDOC.valorOperacionesGratuitas == "0.00")
                    Cadena += string.Format("DEIM|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}",
                    oBeFEDOC.Detalles[i].IgvMontoGratuito == "0.00" ? oBeFEDOC.Detalles[i].IgvMonto : oBeFEDOC.Detalles[i].IgvMontoGratuito,
                    oBeFEDOC.Detalles[i].valorItem != "0.00" ? oBeFEDOC.Detalles[i].valorItem : oBeFEDOC.Detalles[i].valorItemGratutito,
                    oBeFEDOC.Detalles[i].IgvMontoGratuito == "0.00" ? oBeFEDOC.Detalles[i].IgvMonto : oBeFEDOC.Detalles[i].IgvMontoGratuito,
                    oBeFEDOC.Detalles[i].NombreTributo == "IGV" ? (Convert.ToDouble(oBeFEDOC.tasaIgv) * 100).ToString() :
                    oBeFEDOC.Detalles[i].IgvMontoGratuito != "0.00" ? (Convert.ToDouble(oBeFEDOC.tasaIgv) * 100).ToString() : "0",
                    oBeFEDOC.Detalles[i].ValorCodigo,
                    oBeFEDOC.Detalles[i].IgvTipo,
                    "",
                    oBeFEDOC.Detalles[i].CodigoTributo,
                    oBeFEDOC.Detalles[i].NombreTributo,
                    oBeFEDOC.Detalles[i].CodigoTributoInt
                    ) + Environment.NewLine;
                    #endregion

                    #region DEGH: GASTOS E INTERESES HIPOTECARIOS
                    //No aplica
                    #endregion

                    #region DEPA: INFORMACION ADCIIONAL
                    //No aplica ningún tipo
                    #endregion

                    #region DETC: DETRACCION SERVICIO DE TRANSPORTE DE CARGA
                    //No aplica
                    #endregion

                    #region DEVR: DETRACCION SERVICIO DE TRANSPORTE DE CARGA - VALOR DE REFERENCIA
                    //No aplica
                    #endregion
                }
                #endregion

                #region DI: IMPUESTOS GLOBALES
                Cadena += string.Format("DI|{0}|{1}|{2}|{3}|{4}|{5}|",
                          oBeFEDOC.documentoExportacion == "Y" ? "0.00" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesGratuitas) > 0 ? MntImpTransGrat.ToString() :
                          Convert.ToDouble(oBeFEDOC.valorTotalAnticipos) > 0 ? Math.Round(Convert.ToDouble(Globals.SumarCadenas(oBeFEDOC.valorOperacionesGravadas, "-" + oBeFEDOC.valorTotalIGVAnticipos)) * Convert.ToDouble(oBeFEDOC.tasaIgv), 2).ToString() :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesExoneradas) > 0 ? "0.00" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesInafectas) > 0 ? "0.00" : oBeFEDOC.valorIgv,

                          oBeFEDOC.documentoExportacion == "Y" ? "0.00" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesGratuitas) > 0 ? MntImpTransGrat.ToString() :
                          Convert.ToDouble(oBeFEDOC.valorTotalAnticipos) > 0 ? Math.Round(Convert.ToDouble(Globals.SumarCadenas(oBeFEDOC.valorOperacionesGravadas, "-" + oBeFEDOC.valorTotalIGVAnticipos)) * Convert.ToDouble(oBeFEDOC.tasaIgv), 2).ToString() :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesExoneradas) > 0 ? "0.00" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesInafectas) > 0 ? "0.00" : oBeFEDOC.valorIgv,

                          oBeFEDOC.documentoExportacion == "Y" ? "9995" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesGratuitas) > 0 ? "9996" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesGravadas) > 0 ? "1000" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesExoneradas) > 0 ? "9997" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesInafectas) > 0 ? "9998" : "1000",

                          oBeFEDOC.documentoExportacion == "Y" ? "EXP" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesGratuitas) > 0 ? "GRA" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesGravadas) > 0 ? "IGV" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesExoneradas) > 0 ? "EXO" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesInafectas) > 0 ? "INA" : "IGV",

                          oBeFEDOC.documentoExportacion == "Y" ? "FRE" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesGratuitas) > 0 ? "FRE" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesGravadas) > 0 ? "VAT" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesExoneradas) > 0 ? "VAT" :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesInafectas) > 0 ? "FRE" : "VAT",

                          oBeFEDOC.documentoExportacion == "Y" ? oBeFEDOC.valorOperacionesExoneradas :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesGratuitas) > 0 ? Globals.SumarCadenas(oBeFEDOC.valorOperacionesGratuitas, "-" + MntImpTransGrat.ToString()) :
                          //Convert.ToDouble(oBeFEDOC.valorOperacionesGravadas) > 0 ? Utils.SumarCadenas(oBeFEDOC.valorOperacionesGravadas, oBeFEDOC.valorTotalAnticipos) :
                          Convert.ToDouble(oBeFEDOC.valorTotalAnticipos) > 0 ? Math.Round(Convert.ToDouble(Globals.SumarCadenas(oBeFEDOC.valorOperacionesGravadas, "-" + oBeFEDOC.valorTotalIGVAnticipos)), 2).ToString() :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesExoneradas) > 0 ? oBeFEDOC.valorOperacionesExoneradas :
                          Convert.ToDouble(oBeFEDOC.valorOperacionesInafectas) > 0 ? oBeFEDOC.valorOperacionesInafectas : Globals.SumarCadenas(oBeFEDOC.valorOperacionesGravadas, oBeFEDOC.valorTotalAnticipos)
                          ) + Environment.NewLine;
                #endregion

                #region DR: DESCUENTOS, CARGOS GLOBALES Y ANTICIPOS
                if (oBeFEDOC.valorDescuentoGlobal != "0.00")
                {
                    Cadena += string.Format("DR|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}",
                   "FALSE",
                   "00", //Codigo Catalogo 5
                   "DESCUENTO",
                   oBeFEDOC.tipoMoneda,
                   oBeFEDOC.valorDescuentoGlobal,
                   oBeFEDOC.tasaIgv != "0.00" ? "1000" : string.Empty, //Catalogo SUNAT 05
                   oBeFEDOC.tasaIgv != "0.00" ? "IGV" : string.Empty, //Nombre de tributo 
                   oBeFEDOC.tasaIgv != "0.00" ? "VAT" : string.Empty, //Codigo internacional del tributo
                   string.Empty,
                   string.Empty) + Environment.NewLine;
                }
                if (oBeFEDOC.otrosCargos != "0.00")
                {
                    Cadena += string.Format("DR|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}",
                   "TRUE",
                   "50", //Codigo Catalogo 5 
                   "RECARGO",
                   oBeFEDOC.tipoMoneda,
                   oBeFEDOC.otrosCargos,
                   oBeFEDOC.tasaIgv != "0.00" ? "1000" : string.Empty,
                   oBeFEDOC.tasaIgv != "0.00" ? "IGV" : string.Empty,
                   oBeFEDOC.tasaIgv != "0.00" ? "VAT" : string.Empty,
                   string.Empty,
                   string.Empty) + Environment.NewLine;
                }
                //Se envío obligatoriamente cuando hay anticipo
                if (oBeFEDOC.ListAnticipo != null)
                {
                    Cadena += string.Format("DR|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}",
                    "false",
                    "04",
                    "Descuentos globales por anticipos gravados que afectan la base imponible del IGV / IVAP",
                    oBeFEDOC.tipoMoneda,
                    oBeFEDOC.valorTotalAnticipos,
                    "1000",
                    "IGV",
                    "VAT",
                    Math.Round(Convert.ToDouble(oBeFEDOC.valorTotalAnticipos) / Convert.ToDouble(Globals.SumarCadenas(Globals.SumarCadenas(oBeFEDOC.valorTotalAnticipos, oBeFEDOC.valorTotalIGVAnticipos), oBeFEDOC.importeTotal)), 5).ToString(),
                    Globals.SumarCadenas(Globals.SumarCadenas(oBeFEDOC.valorTotalAnticipos, oBeFEDOC.valorTotalIGVAnticipos), oBeFEDOC.importeTotal) + Environment.NewLine);//identificador del pago
                }

                #endregion

                #region SECCION PP: ANTICIPOS Y PREPAGOS
                if (oBeFEDOC.ListAnticipo != null)
                {
                    for (int i = 0; i < oBeFEDOC.ListAnticipo.Count; i++)
                    {
                        Cadena += string.Format("PP|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}",
                        i + 1,
                        //Utils.SumarCadenas(oBeFEDOC.ListAnticipo[i].totalAnticipos, oBeFEDOC.ListAnticipo[i].valorIGV),
                        Globals.SumarCadenas(oBeFEDOC.ListAnticipo[i].totalAnticipos, oBeFEDOC.ListAnticipo[i].valorIGV),
                        //oBeFEDOC.ListAnticipo[i].totalAnticipos,
                        string.Empty, //fecha recepcion
                        string.Empty, //fecha pago
                        string.Empty, //hora pago
                                      //oBeFEDOC.ListAnticipo[i].tipoDocumento == "01" ? "02" : "03",
                        oBeFEDOC.ListAnticipo[i].serieDocumento.Substring(0, 1) == "F" ? "02" : "03",
                        oBeFEDOC.ListAnticipo[i].serieDocumento + "-" + oBeFEDOC.ListAnticipo[i].correlativoDocumento,
                        oBeFEDOC.emisor.tipoIdentificacion,
                        oBeFEDOC.emisor.ruc,
                        i + 1) + Environment.NewLine;//identificador del pago
                    }
                }
                #endregion

                #region SECCION RE: REFERENCIA
                if (oBeFEDOC.tipoComprobante == "01" || oBeFEDOC.tipoComprobante == "03")
                {
                    if (!string.IsNullOrEmpty(oBeFEDOC.guias))
                    {
                        List<string> guias = oBeFEDOC.guias.Split('|').ToList();
                        for (int i = 0; i < guias.Count; i++)
                        {
                            Cadena += string.Format("RE|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}",
                            string.Empty,
                            string.Empty,
                            string.Empty,
                            string.Empty,
                            string.Empty, //No se referencia guías
                            string.Empty,
                            "99",
                            guias[i]) + Environment.NewLine;
                        }
                    }
                }

                if (oBeFEDOC.tipoComprobante == "07" || oBeFEDOC.tipoComprobante == "08")
                {
                    for (int i = 0; i < oBeFEDOC.ListReferencia.Count; i++)
                    {
                        Cadena += string.Format("RE|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}",
                        oBeFEDOC.ListReferencia[i].serie + "-" + oBeFEDOC.ListReferencia[i].correlativo,
                        oBeFEDOC.ListReferencia[i].fechaEmision,
                        oBeFEDOC.ListReferencia[i].tipoDocumento,
                        oBeFEDOC.ListReferencia[i].tipoDocumento == "01" ? "FACTURA" : "BOLETA DE VENTA",
                        string.Empty, //No se referencia guías
                        string.Empty,
                        string.Empty,
                        string.Empty,
                        string.Empty,
                        string.Empty) + Environment.NewLine;
                    }
                }
                #endregion

                #region SECCION FP: Forma de pago
                if (oBeFEDOC.tipoComprobante == "01" || oBeFEDOC.tipoComprobante == "03" || oBeFEDOC.tipoComprobante == "07")
                {
                    //if (oBeFEDOC.notaTipo == "" || oBeFEDOC.notaTipo == "13")
                    {
                        Cadena += string.Format("FP|FormaPago|{0}|{1}|{2}|",
                              oBeFEDOC.condicionPago,
                              oBeFEDOC.condicionPago == "Credito" ? oBeFEDOC.importeTotal : "",
                              string.Empty
                              ) + Environment.NewLine;
                        if (oBeFEDOC.condicionPago == "Credito")
                        {
                            Cadena += string.Format("FP|FormaPago|{0}|{1}|{2}",
                                      "Cuota001",
                                      oBeFEDOC.importeTotal,
                                      oBeFEDOC.fechaVencimiento
                                      ) + Environment.NewLine;
                        }
                    }
                }
                #endregion

                Cadena += "PES|MensajesAt|" + Environment.NewLine;
                Cadena += "PESD|1|Agentes de Retención R.S. 228-2012/SUNAT a partir de 01/11/2012|" + Environment.NewLine;
                Cadena += "PESD|2|(*) Numeral 10 del artículo 33 del TUO de la Ley del IGV e Impuesto Selectivo al Consumo, D.S. N° 055-99-EF y normas modificatorias|" + Environment.NewLine;

                if (oBeFEDOC.TieneDetraccion == "Y")
                    Cadena += "PESD|3|OPERACION SUJETA AL SISTEMA DE PAGOS DE OBLIGACIONES TRIBUTARIAS CON EL GOBIERNO CENTRAL|" + Environment.NewLine;

                Cadena += "]]>";

                try
                {
                    System.IO.File.WriteAllText(Globals.PathBackup + oBeFEDOC.emisor.ruc + "_" + oBeFEDOC.tipoComprobante + "_" + oBeFEDOC.serie + "_" + oBeFEDOC.correlativo + ".TXT", Cadena);
                    Folio = oBeFEDOC.serie + "-" + oBeFEDOC.correlativo;

                    var client = new SftpClient(Globals.SFTP_IP, Globals.SFTP_PORT, Globals.SFTP_USER, Globals.SFTP_PASS);

                    client.Connect();
                    if (client.IsConnected)
                    {
                        using (var fs = new FileStream(Globals.PathBackup + oBeFEDOC.emisor.ruc + "_" + oBeFEDOC.tipoComprobante + "_" + oBeFEDOC.serie + "_" + oBeFEDOC.correlativo + ".TXT", FileMode.Open))
                        {
                            client.UploadFile(fs, Globals.PathArchivos + oBeFEDOC.emisor.ruc + "_" + oBeFEDOC.tipoComprobante + "_" + oBeFEDOC.serie + "_" + oBeFEDOC.correlativo + ".TXT");
                        }
                        EstadoDocumento = "01";
                        GC.Collect();
                        result = true;
                    }
                    else
                    {
                        throw new InvalidOperationException("No se pudo conectar con el servicio SFTP de Paperless");
                    }
                }
                catch (System.Exception ex)
                {
                    mensaje_error = ex.Message;
                    result = false;
                }
            }
            catch (Exception ex)
            {
                mensaje_error = ex.Message;
                result = false;
            }
            return result;
        }

        public static bool RegistrarBaja(Baja document, ref string estadoDocumento, ref string mensajeError)
        {
            bool result = false;
            try
            {
                var client = new IntegradorEstelaBaja.WssCargaBajasSoapClient();

                Globals.ValidaCamposBaja(document);

                var listaBajas = new ArrayOfBajas();
                var baja = new bajas
                {
                    Tipo_Docu = document.tipoComprobante,
                    Serie_Inte = document.serie,
                    Foli_Inte = document.correlativo,
                    Fech_Emis = document.fechaEmision,
                    Motiv_Anul = document.notaSustento,
                };
                listaBajas.Add(baja);

                var response = client.cargaBajas(Globals.oOADM.TaxIdNum, listaBajas);
                List<string> rpta = response.ToString().Split(';').ToList();
                client.Close();

                string mensaje = string.Empty;
                switch (response.Codigo)
                {
                    case "DOK":
                        estadoDocumento = "04";
                        mensajeError = "El documento " + Globals.oOADM.TaxIdNum + "-" + document.tipoComprobante + "-" + document.serie + "-" + document.correlativo + " ha sido dado de baja correctamente.";
                        break;
                    default:
                        estadoDocumento = document.Estado;
                        mensajeError = response.Mensajes;
                        break;
                }

            }
            catch (Exception ex)
            {
                estadoDocumento = document.Estado;
                mensajeError = ex.Message;
            }
            return result;
        }

        public static void ConsultarEstadoCPE(string tipoDoc, List<Comprobante> listaComp)
        {
            try
            {
                switch (tipoDoc)
                {
                    case "01":
                    case "03":
                    case "07":
                    case "08":
                        foreach (Comprobante document in listaComp)
                        {
                            string mensajeerror = "";
                            string nombreArchivo = "";
                            string pdf_url = "";
                            string xml_url = "";
                            string cdr_url = "";
                            bool respuestaanexo = false;

                            Entities.Paperless.Documento oBeFEDOC = new Entities.Paperless.Documento();
                            System.Data.DataTable dtDetalles = new System.Data.DataTable();
                            DateTime date1 = DateTime.Now;

                            oBeFEDOC.DocEntry = document.DocEntry.ToString();

                            SAPDI.ObtenerDatosCPEPaperless(ref oBeFEDOC, document.Tabla);

                            BasicHttpBinding binding = new BasicHttpBinding();
                            binding.Security.Mode = BasicHttpSecurityMode.None;  // Cambiar si el servicio usa seguridad
                            binding.MaxReceivedMessageSize = 2000000;

                            // Dirección del endpoint
                            EndpointAddress endpoint = new EndpointAddress(Globals.URL_OSE);
                            // Creación del cliente del servicio
                            ChannelFactory<OnlinePortType> factory = new ChannelFactory<OnlinePortType>(binding, endpoint);
                            OnlinePortType client = factory.CreateChannel();
                            nombreArchivo = oBeFEDOC.emisor.ruc + "_" + oBeFEDOC.tipoComprobante + "_" + oBeFEDOC.serie + "_" + oBeFEDOC.correlativo;

                            //Obtener estado
                            // Creación del request
                            var requestRpta = new OnlineRecoveryRequest
                            {
                                ruc = Globals.oOADM.TaxIdNum,
                                login = "adminppl",
                                clave = "abc123",
                                tipoDoc = Convert.ToInt32(oBeFEDOC.tipoComprobante),
                                folio = oBeFEDOC.serie + "-" + oBeFEDOC.correlativo,
                                tipoRetorno = 3 //Obtiene PDF
                            };

                            var responseRpta = client.OnlineRecovery(requestRpta);

                            if (responseRpta != null)
                            {
                                var xmlResponseEstado = responseRpta.@return.ToString().Replace("&", "&amp;");

                                var docEstado = new XmlDocument();
                                docEstado.LoadXml(xmlResponseEstado);

                                var codigoEstado = docEstado.SelectSingleNode("//Codigo")?.InnerText;
                                var mensajeEstado = docEstado.SelectSingleNode("//Mensaje")?.InnerText;

                                switch (codigoEstado)
                                {
                                    case "0":
                                        SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, "En espera: " + mensajeEstado, "01", Globals.oCompany);
                                        break;
                                    case "1":
                                    case "2":
                                        var request = new OnlineRecoveryRequest
                                        {
                                            ruc = Globals.oOADM.TaxIdNum,
                                            login = "adminppl",
                                            clave = "abc123",
                                            tipoDoc = Convert.ToInt32(oBeFEDOC.tipoComprobante),
                                            folio = oBeFEDOC.serie + "-" + oBeFEDOC.correlativo,
                                            tipoRetorno = 2 //Obtiene PDF
                                        };

                                        //Obtiene PDF
                                        var response = client.OnlineRecovery(request);

                                        // Procesar la respuesta
                                        if (response != null)
                                        {
                                            var xmlResponse = response.@return.ToString().Replace("&", "&amp;");

                                            var doc = new XmlDocument();
                                            doc.LoadXml(xmlResponse);

                                            var codigo = doc.SelectSingleNode("//Codigo")?.InnerText;
                                            var mensaje = doc.SelectSingleNode("//Mensaje")?.InnerText;
                                            if (mensaje != "No se encontro respuesta SUNAT")
                                            {
                                                if (codigo == "0" && !string.IsNullOrEmpty(mensaje))
                                                    pdf_url = mensaje;
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine("Obterner PDF:No se recibió respuesta del servicio.");
                                        }

                                        //Obtiene XML
                                        var request2 = new OnlineRecoveryRequest
                                        {
                                            ruc = Globals.oOADM.TaxIdNum,
                                            login = "adminppl",
                                            clave = "abc123",
                                            tipoDoc = Convert.ToInt32(oBeFEDOC.tipoComprobante),
                                            folio = oBeFEDOC.serie + "-" + oBeFEDOC.correlativo,
                                            tipoRetorno = 1 //Obtiene XML
                                        };

                                        var response2 = client.OnlineRecovery(request2);
                                        if (response2 != null)
                                        {
                                            var xmlResponse = response2.@return.ToString().Replace("&", "&amp;");

                                            var doc = new XmlDocument();
                                            doc.LoadXml(xmlResponse);

                                            var codigo = doc.SelectSingleNode("//Codigo")?.InnerText;
                                            var mensaje = doc.SelectSingleNode("//Mensaje")?.InnerText;

                                            if (mensaje != "No se encontro respuesta SUNAT")
                                            {
                                                if (codigo == "0" && !string.IsNullOrEmpty(mensaje))
                                                    xml_url = mensaje;
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine("Obterner XML:No se recibió respuesta del servicio.");
                                        }

                                        //Obtiene CDR
                                        var request3 = new OnlineRecoveryRequest
                                        {
                                            ruc = Globals.oOADM.TaxIdNum,
                                            login = "adminppl",
                                            clave = "abc123",
                                            tipoDoc = Convert.ToInt32(oBeFEDOC.tipoComprobante),
                                            folio = oBeFEDOC.serie + "-" + oBeFEDOC.correlativo,
                                            tipoRetorno = 8
                                        };

                                        var response3 = client.OnlineRecovery(request3);
                                        if (response3 != null)
                                        {
                                            var xmlResponse = response3.@return.ToString().Replace("&", "&amp;");

                                            var doc = new XmlDocument();

                                            if (!string.IsNullOrEmpty(xmlResponse))
                                            {
                                                doc.LoadXml(xmlResponse);

                                                var codigo = doc.SelectSingleNode("//Codigo")?.InnerText;
                                                var mensaje = doc.SelectSingleNode("//Mensaje")?.InnerText;

                                                if (mensaje != "No se encontro respuesta SUNAT")
                                                {
                                                    if (codigo == "0" && !string.IsNullOrEmpty(mensaje))
                                                        cdr_url = mensaje;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine("Obterner XML:No se recibió respuesta del servicio.");
                                        }

                                        // Cerrar el cliente
                                        ((IClientChannel)client).Close();

                                        bool tieneCdr = true;
                                        switch (tipoDoc)
                                        {
                                            case "01":
                                                if (string.IsNullOrEmpty(cdr_url))
                                                    tieneCdr = false;
                                                break;
                                            case "03":
                                                tieneCdr = true;
                                                break;
                                            case "07":
                                            case "08":
                                                if (oBeFEDOC.ListReferencia[0].tipoDocumento == "01" && string.IsNullOrEmpty(cdr_url))
                                                    tieneCdr = false;
                                                break;
                                        }

                                        if (!string.IsNullOrEmpty(pdf_url) && !string.IsNullOrEmpty(xml_url) && tieneCdr)
                                        {
                                            respuestaanexo = SAPDI.AnexarDocumentosCpe(document.DocEntry, nombreArchivo, pdf_url, xml_url, cdr_url, document.ObjType, Globals.oCompany, ref mensajeerror);
                                            SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, mensajeEstado, "02", Globals.oCompany);
                                        }
                                        break;
                                    case "3":
                                        SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, mensajeEstado, "03", Globals.oCompany);
                                        break;
                                    default:
                                        SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, mensajeEstado, "01", Globals.oCompany);
                                        break;
                                }
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
            }
        }

        public static void ValidaCamposCPE(Entities.Paperless.Documento oBeFEDOC)
        {
            try
            {
                if (string.IsNullOrEmpty(oBeFEDOC.tipoOperacion))
                    throw new Exception("Debe seleccionar un tipo de operación");

                if (oBeFEDOC.tipoComprobante == "07")
                {
                    if (string.IsNullOrEmpty(oBeFEDOC.notaTipo))
                        throw new Exception("Debe seleccionar un tipo de nota de crédito");
                    if (string.IsNullOrEmpty(oBeFEDOC.notaSustento))
                        throw new Exception("Debe indicar un sustento para la nota de crédito");
                }

                if (oBeFEDOC.tipoComprobante == "08")
                {
                    if (string.IsNullOrEmpty(oBeFEDOC.notaTipo))
                        throw new Exception("Debe seleccionar un tipo de nota de débito");
                    if (string.IsNullOrEmpty(oBeFEDOC.notaSustento))
                        throw new Exception("Debe indicar un sustento para la nota de débito");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
