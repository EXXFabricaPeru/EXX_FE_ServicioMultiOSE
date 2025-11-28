using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ServicioConectorFE.Core;
using ServicioConectorFE.Entities.Nubefact;
using ServicioConectorFE.Entities.SAP;
using ServicioConectorFE.Framework;
using ServicioConectorFE.IntegradorEstelaAdjunto;
using ServicioConectorFE.IntegradorEstelaBaja;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Security.Tokens;
using System.Text;
using System.Threading.Tasks;

namespace ServicioConectorFE.OSE
{
    public class Nubefact
    {
        public static bool RegistrarRetencion(string docmun, int objectype, string tipodocumento, string Tabla, int docentry, ref string EstadoDocumento, ref string Folio, ref string mensaje_error)
        {
            bool result = false;

            Entities.Estela.Retencion oDocumentoRetFE = new Entities.Estela.Retencion();
            Entities.Nubefact.Retencion oRet = new Entities.Nubefact.Retencion();

            try
            {
                DateTime date1 = DateTime.Now;
                SAPDI.ObtenerDatosCRE(docentry, ref oDocumentoRetFE);

                Globals.ValidaCamposCRE(oDocumentoRetFE);

                oRet.Operacion = "generar_retencion";
                oRet.Serie = "RRR1"; //oDocumentoRetFE.serieNumeroRetencion.Split('-').ToList()[0];
                oRet.Numero = Convert.ToInt32(oDocumentoRetFE.serieNumeroRetencion.Split('-').ToList()[1]);
                oRet.FechaDeEmision = Convert.ToDateTime(oDocumentoRetFE.fechaEmision).ToString("dd-MM-yyyy");
                oRet.ClienteTipoDeDocumento = Convert.ToInt32(oDocumentoRetFE.tipoDocumentoProveedor);
                oRet.ClienteNumeroDeDocumento = oDocumentoRetFE.numeroDocumentoProveedor;
                oRet.ClienteDenominacion = oDocumentoRetFE.razonSocialEmisor;
                oRet.ClienteDireccion = oDocumentoRetFE.direccionEmisor.Replace('\r', ' ').Replace(Environment.NewLine, "");

                var email = oDocumentoRetFE.correoAdquiriente.Split(';').ToList();
                if (email.Count > 0)
                {
                    oRet.ClienteEmail = email[0].Trim();
                    if (email.Count > 1) oRet.ClienteEmail = email[1].Trim();
                    if (email.Count > 2) oRet.ClienteEmail = email[2].Trim();
                }

                if (oDocumentoRetFE.tipoMonedaTotalPagado == "SOL" || oDocumentoRetFE.tipoMonedaTotalPagado == "PEN")
                {
                    oRet.Moneda = "1";
                    oRet.TipoDeTasaDeRetencion = Convert.ToDouble(oDocumentoRetFE.tasaRetencion) == 6 ? "2" : "1";
                    oRet.TotalRetenido = oDocumentoRetFE.importeTotalRetenido;
                    oRet.TotalPagado = oDocumentoRetFE.importeTotalPagado;
                }
                else
                {
                    oRet.Moneda = "1";
                    oRet.TipoDeTasaDeRetencion = Convert.ToDouble(oDocumentoRetFE.tasaRetencion) == 6 ? "2" : "1";
                    oRet.TotalRetenido = oDocumentoRetFE.importeTotalRetenido;
                    oRet.TotalPagado = oDocumentoRetFE.importeTotalPagado;
                }

                oRet.Observaciones = oDocumentoRetFE.observaciones;
                oRet.EnviarAutomaticamenteALaSunat = "true";
                oRet.EnviarAutomaticamenteAlCliente = "false";
                oRet.CodigoUnico = "";
                oRet.FormatoDePdf = "A4";

                // Inicializar la lista de items
                oRet.Items = new List<RetencionItem>();

                for (int i = 0; i < oDocumentoRetFE.RetencionItems.Count; i++)
                {
                    RetencionItem item = new RetencionItem();
                    item.DocumentoRelacionadoTipo = oDocumentoRetFE.RetencionItems[i].tipoDocumentoRelacionado;
                    item.DocumentoRelacionadoSerie = oDocumentoRetFE.RetencionItems[i].numeroDocumentoRelacionado.Split('-').ToList()[0];
                    item.DocumentoRelacionadoNumero = Convert.ToInt32(oDocumentoRetFE.RetencionItems[i].numeroDocumentoRelacionado.Split('-').ToList()[1]);
                    item.DocumentoRelacionadoFechaDeEmision = Convert.ToDateTime(oDocumentoRetFE.RetencionItems[i].fechaEmisionDocumentoRelacionado).ToString("dd-MM-yyyy");
                    item.DocumentoRelacionadoMoneda = oDocumentoRetFE.RetencionItems[i].tipoMonedaDocumentoRelacionado == "SOL" || oDocumentoRetFE.RetencionItems[i].tipoMonedaDocumentoRelacionado == "PEN" ? "1" : "2";
                    item.PagoFecha = Convert.ToDateTime(oDocumentoRetFE.RetencionItems[i].fechaPago).ToString("dd-MM-yyyy");
                    item.PagoNumero = oDocumentoRetFE.RetencionItems[i].numeroPago;
                    item.ImporteRetenidoFecha = Convert.ToDateTime(oDocumentoRetFE.RetencionItems[i].fechaEmisionDocumentoRelacionado).ToString("dd-MM-yyyy");

                    if (item.DocumentoRelacionadoMoneda == "1")
                    {
                        item.DocumentoRelacionadoTotal = oDocumentoRetFE.RetencionItems[i].importeTotalDocumentoRelacionado;
                        item.PagoTotalSinRetencion = oDocumentoRetFE.RetencionItems[i].importePagoSinRetencion;
                        item.ImporteRetenido = oDocumentoRetFE.RetencionItems[i].importeRetenido;
                        item.ImportePagadoConRetencion = Globals.SumarCadenas(oDocumentoRetFE.RetencionItems[i].importeTotalPagarNeto, oDocumentoRetFE.RetencionItems[i].importeRetenido);
                    }
                    else
                    {
                        item.TipoDeCambioFecha = Convert.ToDateTime(oDocumentoRetFE.RetencionItems[i].fechaCambio).ToString("dd-MM-yyyy");
                        item.TipoDeCambio = oDocumentoRetFE.RetencionItems[i].factorTipoCambioMoneda;
                        item.ImporteRetenido = Math.Round(Convert.ToDouble(oDocumentoRetFE.RetencionItems[i].importeRetenido) / Convert.ToDouble(item.TipoDeCambio), 2, MidpointRounding.AwayFromZero).ToString();
                        item.DocumentoRelacionadoTotal = oDocumentoRetFE.RetencionItems[i].importeTotalDocumentoRelacionado;
                        item.PagoTotalSinRetencion = Globals.SumarCadenas(oDocumentoRetFE.RetencionItems[i].importePagoSinRetencion, "-" + item.ImporteRetenido);
                        item.ImportePagadoConRetencion = Math.Round(Convert.ToDouble(Globals.SumarCadenas(oDocumentoRetFE.RetencionItems[i].importeTotalPagarNeto, oDocumentoRetFE.RetencionItems[i].importeRetenido)) / Convert.ToDouble(item.TipoDeCambio), 2, MidpointRounding.AwayFromZero).ToString();
                    }
                    oRet.Items.Add(item);
                }

                try
                {
                    string json = JsonConvert.SerializeObject(oRet, Formatting.Indented);
                    Console.WriteLine(json);
                    byte[] bytes = Encoding.Default.GetBytes(json);
                    string json_en_utf_8 = Encoding.UTF8.GetString(bytes);

                    string json_de_respuesta = SendJson(Globals.URL_OSE, json_en_utf_8, Globals.TOKEN);
                    dynamic r = JsonConvert.DeserializeObject<Respuesta>(json_de_respuesta);
                    string r2 = JsonConvert.SerializeObject(r, Formatting.Indented);
                    dynamic json_r_in = JsonConvert.DeserializeObject<Respuesta>(r2);

                    dynamic leer_respuesta = JsonConvert.DeserializeObject<Respuesta>(json_de_respuesta);
                    if (leer_respuesta.errors == null)
                    {
                        Folio = oDocumentoRetFE.serieNumeroRetencion;
                        EstadoDocumento = "01";
                        result = true;
                    }
                    else
                    {
                        if (leer_respuesta.errors != "Este documento ya existe en NubeFacT")
                            throw new Exception(leer_respuesta.errors);
                        else
                        {
                            Folio = oDocumentoRetFE.serieNumeroRetencion;
                            EstadoDocumento = "01";
                            result = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            catch (Exception ex)
            {
                mensaje_error = ex.Message;
                return false;
            }
            finally
            {
                GC.Collect();
            }
            return result;
        }

        public static bool RegistrarGuia(string docmun, int objectype, string tipodocumento, string Tabla, int docentry, ref string EstadoDocumento, ref string Folio, ref string mensaje_error)
        {
            bool result = false;
            Entities.Estela.Documento oBeFEDOC = new Entities.Estela.Documento();
            System.Data.DataTable dtDetalles = new System.Data.DataTable();

            try
            {
                DateTime date1 = DateTime.Now;

                oBeFEDOC.DocEntry = docentry.ToString();
                SAPDI.ObtenerDatosCPEEstela(ref oBeFEDOC, Tabla);
                oBeFEDOC.Detalles = new List<Entities.Estela.DocLine>();
                SAPDI.ObtenerDetalleDocumento(ref dtDetalles, Tabla, docentry);

                String Cadena = string.Empty;

            }
            catch (Exception ex)
            {
                mensaje_error = ex.Message;
                return false;
            }
            return result;
        }

        public static bool RegistrarDocumento(string docmun, int objectype, string tipodocumento, string Tabla, int docentry, ref string EstadoDocumento, ref string Folio, ref string mensaje_error)
        {
            bool result = false;
            Entities.Estela.Documento oBeFEDOC = new Entities.Estela.Documento();
            Entities.Nubefact.Documento oBeFE = new Entities.Nubefact.Documento();
            System.Data.DataTable dtDetalles = new System.Data.DataTable();

            try
            {
                DateTime date1 = DateTime.Now;

                oBeFEDOC.DocEntry = docentry.ToString();
                SAPDI.ObtenerDatosCPEEstela(ref oBeFEDOC, Tabla);
                oBeFEDOC.Detalles = new List<Entities.Estela.DocLine>();
                SAPDI.ObtenerDetalleDocumento(ref dtDetalles, Tabla, docentry);

                String Cadena = string.Empty;

                Globals.ValidaCamposCPE(oBeFEDOC);

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


                oBeFEDOC.Detalles = new List<Entities.Estela.DocLine>();
                int RowsinMatrix = 1;

                if (oBeFEDOC.DocType == "I") //Articulos
                {
                    foreach (System.Data.DataRow detalle in dtDetalles.Rows)
                    {
                        Entities.Estela.DocLine oDetalle = new Entities.Estela.DocLine();
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
                        oDetalle.IgvTipo = detalle.Field<string>("U_EXX_TIPOAFECT_FE").ToString();
                        oDetalle.descuentoPct = Globals.RetirarMonedayDecimal(detalle.Field<string>("DiscPrcnt"));
                        oDetalle.OperacionAfecta = oDetalle.CodigoImpuesto == "IGV" ? "A" : oDetalle.CodigoImpuesto == "INA" ? "I" : oDetalle.CodigoImpuesto == "EXO" ? "E" : ""; // detalle.Field<string>("U_BPP_OPER").ToString();
                        oDetalle.TaxOnly = detalle.Field<string>("TaxOnly").ToString();
                        //oDetalle.Glosa = detalle.Field<string>("U_EXA_GLOSA").ToString();

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
                            if (string.IsNullOrEmpty(oDetalle.IgvTipo))
                            {
                                mensaje_error = "Para transferencia gratuita debe seleccionar el campo FE TIPO OPERACION en los detalles del documento";
                                return false;
                            }
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
                                oDetalle.IgvTipo = "4";
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
                                    oDetalle.IgvTipo = "2";
                                    MntExonerado += Convert.ToDouble(oDetalle.valorItem);
                                }
                                else if (oDetalle.OperacionAfecta == "I")
                                {
                                    oDetalle.ValorCodigo = "O";
                                    oDetalle.CodigoTributo = "9998";
                                    oDetalle.CodigoTributoInt = "FRE";
                                    oDetalle.NombreTributo = "INA";
                                    oDetalle.IgvTipo = "3";
                                    MntInafecto += Convert.ToDouble(oDetalle.valorItem);
                                }
                                else if (oDetalle.OperacionAfecta == "A")
                                {
                                    oDetalle.ValorCodigo = "S";
                                    oDetalle.CodigoTributo = "1000";
                                    oDetalle.CodigoTributoInt = "VAT";
                                    oDetalle.NombreTributo = "IGV";
                                    oDetalle.IgvTipo = "1";
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
                        Entities.Estela.DocLine oDetalle = new Entities.Estela.DocLine();
                        oDetalle.posicion = RowsinMatrix.ToString();
                        oDetalle.codigo = detalle.Field<string>("descripcionServ").ToString();
                        oDetalle.descripcion = detalle.Field<string>("descripcion").ToString();
                        oDetalle.unidadMedidaSAP = detalle.Field<string>("unidadMedidaSAP").ToString();
                        oDetalle.unidadMedida = "NIU";
                        oDetalle.cantidadUnidades = "1.00";
                        oDetalle.unidadMedidaSAP = detalle.Field<string>("unidadMedidaSAP").ToString();
                        oDetalle.unidadMedida = detalle.Field<string>("unidadMedida").ToString();
                        oDetalle.valorUnitario = Globals.RetirarMonedayDecimal(detalle.Field<string>("valorUnitario").ToString(), 6);
                        oDetalle.IgvMonto = Globals.RetirarMonedayDecimal(detalle.Field<string>("montoIGV").ToString());
                        oDetalle.precioUnitario = Globals.RetirarMonedayDecimal(detalle.Field<string>("precioUnitario").ToString());
                        oDetalle.valorItem = Globals.RetirarMonedayDecimal(detalle.Field<string>("valorItem").ToString());
                        oDetalle.CodigoImpuesto = detalle.Field<string>("CodigoImpuesto").ToString();
                        oDetalle.tasaIGV = Convert.ToDouble(detalle.Field<string>("tasaIGV"));
                        oDetalle.Rate = Globals.RetirarMonedayDecimal(detalle.Field<string>("Rate").ToString());
                        oDetalle.IgvTipo = detalle.Field<string>("U_EXX_TIPOAFECT_FE").ToString();
                        oDetalle.descuentoPct = Globals.RetirarMonedayDecimal(detalle.Field<string>("DiscPrcnt"));
                        oDetalle.OperacionAfecta = oDetalle.CodigoImpuesto == "IGV" ? "A" : oDetalle.CodigoImpuesto == "INA" ? "I" : oDetalle.CodigoImpuesto == "EXO" ? "E" : ""; // detalle.Field<string>("U_BPP_OPER").ToString();
                        oDetalle.TaxOnly = detalle.Field<string>("TaxOnly").ToString();
                        //oDetalle.Glosa = detalle.Field<string>("U_EXA_GLOSA").ToString();

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
                            if (string.IsNullOrEmpty(oDetalle.IgvTipo))
                            {
                                mensaje_error = "Para transferencia gratuita debe seleccionar el campo FE TIPO OPERACION en los detalles del documento";
                                return false;
                            }
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
                                oDetalle.IgvTipo = "4";
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
                                    oDetalle.IgvTipo = "2";
                                    MntExonerado += Convert.ToDouble(oDetalle.valorItem);
                                }
                                else if (oDetalle.OperacionAfecta == "I")
                                {
                                    oDetalle.ValorCodigo = "O";
                                    oDetalle.CodigoTributo = "9998";
                                    oDetalle.CodigoTributoInt = "FRE";
                                    oDetalle.NombreTributo = "INA";
                                    oDetalle.IgvTipo = "3";
                                    MntInafecto += Convert.ToDouble(oDetalle.valorItem);
                                }
                                else if (oDetalle.OperacionAfecta == "A")
                                {
                                    oDetalle.ValorCodigo = "S";
                                    oDetalle.CodigoTributo = "1000";
                                    oDetalle.CodigoTributoInt = "VAT";
                                    oDetalle.NombreTributo = "IGV";
                                    oDetalle.IgvTipo = "1";
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
                        Globals.ObtenerAnticiposEstela(ref oBeFEDOC, Tabla, Convert.ToInt32(oBeFEDOC.DocEntry));
                }

                MntTotal = MntGrabado + MntInafecto + MntExonerado + MntIGV - Convert.ToDouble(Globals.SumarCadenas(oBeFEDOC.valorTotalAnticipos, oBeFEDOC.valorTotalIGVAnticipos));

                oBeFEDOC.valorOperacionesGravadas = Globals.RetirarMonedayDecimal(MntGrabado.ToString());
                oBeFEDOC.valorOperacionesExoneradas = Globals.RetirarMonedayDecimal(MntExonerado.ToString());
                oBeFEDOC.valorOperacionesInafectas = Globals.RetirarMonedayDecimal(MntInafecto.ToString());
                oBeFEDOC.valorOperacionesGratuitas = Globals.RetirarMonedayDecimal(MntTransGratuita.ToString());
                oBeFEDOC.importeTotal = Globals.RetirarMonedayDecimal(MntTotal.ToString());
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
                oBeFEDOC.valorIgv = Convert.ToDouble(oBeFEDOC.valorIgv) < 0 ? "0.00" : oBeFEDOC.valorIgv;


                //Armado de json nubefact
                oBeFE.Operacion = "generar_comprobante";
                oBeFE.TipoDeComprobante = oBeFEDOC.tipoComprobante == "01" ? "1" : oBeFEDOC.tipoComprobante == "03" ? "2" : oBeFEDOC.tipoComprobante == "07" ? "3" : "4";
                oBeFE.Serie = "FFF1"; // oBeFEDOC.serie;
                oBeFE.Numero = Convert.ToInt32(oBeFEDOC.correlativo);
                oBeFE.SunatTransaction = 1;
                oBeFE.ClienteTipoDeDocumento = Convert.ToInt32(oBeFEDOC.adquiriente.tipoIdentificacion);
                oBeFE.ClienteNumeroDeDocumento = oBeFEDOC.adquiriente.numeroIdentificacion;
                oBeFE.ClienteDenominacion = oBeFEDOC.adquiriente.nombre;
                oBeFE.ClienteDireccion = oBeFEDOC.adquiriente.direccionFiscal;
                var email = oBeFEDOC.adquiriente.email.Split(';').ToList();
                if (email.Count > 0)
                {
                    oBeFE.ClienteEmail = email[0].Trim();
                    if (email.Count > 1) oBeFE.ClienteEmail = email[1].Trim();
                    if (email.Count > 2) oBeFE.ClienteEmail = email[2].Trim();
                }

                oBeFE.FechaDeEmision = DateTime.Now.ToString("dd-MM-yyyy"); //Convert.ToDateTime(oBeFEDOC.fechaEmision).ToString("dd-MM-yyyy");
                oBeFE.FechaDeVencimiento = Convert.ToDateTime(oBeFEDOC.fechaVencimiento).ToString("dd-MM-yyyy");
                oBeFE.Moneda = oBeFEDOC.tipoMoneda == "SOL" || oBeFEDOC.tipoMoneda == "PEN" ? "1" : "2";
                if (oBeFE.Moneda == "1") oBeFE.TipoDeCambio = "";
                else oBeFE.TipoDeCambio = oBeFEDOC.TipoCambio;
                oBeFE.PorcentajeDeIgv = Convert.ToDouble(oBeFEDOC.tasaIgv) != 0 ? (Convert.ToDouble(oBeFEDOC.tasaIgv) * 100).ToString() : "0";

                oBeFE.DescuentoGlobal = string.IsNullOrEmpty(oBeFEDOC.varlorDescuentoGlobal) ? "0" : oBeFEDOC.varlorDescuentoGlobal;
                oBeFE.TotalDescuento = string.IsNullOrEmpty(oBeFEDOC.varlorDescuentoGlobal) ? "0" : oBeFEDOC.varlorDescuentoGlobal;
                oBeFE.TotalAnticipo = string.IsNullOrEmpty(oBeFEDOC.valorTotalAnticipos) ? "0" : oBeFEDOC.valorTotalAnticipos;
                oBeFE.TotalGravada = string.IsNullOrEmpty(oBeFEDOC.valorOperacionesGravadas) ? "0" : oBeFEDOC.valorOperacionesGravadas;
                oBeFE.TotalInafecta = string.IsNullOrEmpty(oBeFEDOC.valorOperacionesInafectas) ? "0" : oBeFEDOC.valorOperacionesInafectas;
                oBeFE.TotalExonerada = string.IsNullOrEmpty(oBeFEDOC.valorOperacionesExoneradas) ? "0" : oBeFEDOC.valorOperacionesExoneradas;
                oBeFE.TotalIgv = string.IsNullOrEmpty(oBeFEDOC.valorIgv) ? "0" : oBeFEDOC.valorIgv;
                oBeFE.TotalGratuita = string.IsNullOrEmpty(oBeFEDOC.valorOperacionesGratuitas) ? "0" : oBeFEDOC.valorOperacionesGratuitas;
                oBeFE.TotalOtrosCargos = string.IsNullOrEmpty(oBeFEDOC.otrosCargos) ? "0" : oBeFEDOC.otrosCargos;
                oBeFE.Total = string.IsNullOrEmpty(oBeFEDOC.importeTotal) ? "0" : oBeFEDOC.importeTotal;

                oBeFE.PercepcionTipo = "";
                oBeFE.PercepcionBaseImponible = "";
                oBeFE.TotalPercepcion = "";
                oBeFE.TotalIncluidoPercepcion = "";


                if (!string.IsNullOrEmpty(oBeFEDOC.codigoDetraccion))
                {
                    oBeFE.Detraccion = "true";
                    oBeFE.detracciontipo = oBeFEDOC.codigoDetraccion;
                    oBeFE.detracciontotal = oBeFEDOC.totalDetraccion;
                }
                else
                {
                    oBeFE.Detraccion = "false";
                    oBeFE.detracciontipo = "";
                    oBeFE.detracciontotal = "";
                }

                oBeFE.Observaciones = oBeFEDOC.observaciones;

                if (oBeFEDOC.tipoComprobante == "07" || oBeFEDOC.tipoComprobante == "08")
                {
                    oBeFE.DocumentoQueSeModificaTipo = oBeFEDOC.ListReferencia[0].tipoDocumento == "01" ? "1" : "2";
                    oBeFE.DocumentoQueSeModificaSerie = oBeFEDOC.ListReferencia[0].serie;
                    oBeFE.DocumentoQueSeModificaNumero = oBeFEDOC.ListReferencia[0].correlativo;
                    if (oBeFEDOC.tipoComprobante == "07")
                    {
                        oBeFE.TipoDeNotaDeCredito = oBeFEDOC.notaTipo;
                        oBeFE.TipoDeNotaDeDebito = "";
                    }
                    else
                    {
                        oBeFE.TipoDeNotaDeCredito = "";
                        oBeFE.TipoDeNotaDeDebito = oBeFEDOC.notaTipo;
                    }
                }
                else
                {
                    oBeFE.DocumentoQueSeModificaTipo = "";
                    oBeFE.DocumentoQueSeModificaSerie = "";
                    oBeFE.DocumentoQueSeModificaNumero = "";
                    oBeFE.TipoDeNotaDeCredito = "";
                    oBeFE.TipoDeNotaDeDebito = "";
                }

                oBeFE.EnviarAutomaticamenteALaSunat = "true";
                oBeFE.EnviarAutomaticamenteAlCliente = "false";
                oBeFE.CodigoUnico = "";
                oBeFE.CondicionesDePago = oBeFEDOC.credito == "Y" ? "Credito" : "Contado";
                oBeFE.MedioDePago = "";
                oBeFE.PlacaVehiculo = "";
                oBeFE.OrdenCompraServicio = oBeFEDOC.ordenCompra;
                oBeFE.TablaPersonalizadaCodigo = "";
                oBeFE.FormatoDePdf = "A4";
                oBeFE.Items = new List<Entities.Nubefact.Item>();

                for (int i = 0; i < oBeFEDOC.Detalles.Count; i++)
                {
                    Entities.Nubefact.Item item = new Entities.Nubefact.Item();
                    item.UnidadDeMedida = oBeFEDOC.Detalles[i].unidadMedida;
                    item.Codigo = oBeFEDOC.Detalles[i].codigo;
                    item.Descripcion = oBeFEDOC.Detalles[i].descripcion;
                    item.Cantidad = oBeFEDOC.Detalles[i].cantidadUnidades;
                    item.ValorUnitario = oBeFEDOC.Detalles[i].valorUnitario;
                    item.PrecioUnitario = oBeFEDOC.Detalles[i].precioUnitario;
                    item.Descuento = oBeFEDOC.Detalles[i].descuento;
                    item.Subtotal = oBeFEDOC.Detalles[i].valorItem;
                    item.TipoDeIgv = oBeFEDOC.Detalles[i].IgvTipo;
                    item.Igv = oBeFEDOC.Detalles[i].IgvMonto;
                    item.Total = Globals.SumarCadenas(oBeFEDOC.Detalles[i].valorItem, oBeFEDOC.Detalles[i].IgvMonto);
                    item.AnticipoRegularizacion = "false";
                    item.AnticipoDocumentoSerie = "";
                    item.AnticipoDocumentoNumero = "";
                    //item.CodigoProductoSunat = oBeFEDOC.Detalles[i].codigo;
                    oBeFE.Items.Add(item);
                }

                try
                {
                    string json = JsonConvert.SerializeObject(oBeFE, Formatting.Indented);
                    Console.WriteLine(json);
                    byte[] bytes = Encoding.Default.GetBytes(json);
                    string json_en_utf_8 = Encoding.UTF8.GetString(bytes);

                    string json_de_respuesta = SendJson(Globals.URL_OSE, json_en_utf_8, Globals.TOKEN);
                    dynamic r = JsonConvert.DeserializeObject<Respuesta>(json_de_respuesta);
                    string r2 = JsonConvert.SerializeObject(r, Formatting.Indented);
                    dynamic json_r_in = JsonConvert.DeserializeObject<Respuesta>(r2);

                    dynamic leer_respuesta = JsonConvert.DeserializeObject<Respuesta>(json_de_respuesta);
                    if (leer_respuesta.errors == null)
                    {
                        Console.WriteLine(json_r_in);
                        Console.WriteLine();
                        Console.WriteLine();
                        Console.WriteLine("TIPO: " + leer_respuesta.tipo);
                        Console.WriteLine("SERIE: " + leer_respuesta.serie);
                        Console.WriteLine("NUMERO: " + leer_respuesta.numero);
                        Console.WriteLine("URL: " + leer_respuesta.url);
                        Console.WriteLine("ACEPTADA_POR_SUNAT: " + leer_respuesta.aceptada_por_sunat);
                        Console.WriteLine("DESCRIPCION SUNAT: " + leer_respuesta.sunat_description);
                        Console.WriteLine("NOTA SUNAT: " + leer_respuesta.sunat_note);
                        Console.WriteLine("CODIGO RESPUESTA SUNAT: " + leer_respuesta.sunat_responsecode);
                        Console.WriteLine("SUNAT ERROR SOAP: " + leer_respuesta.sunat_soap_error);
                        Console.WriteLine("PDF EN BASE64: " + leer_respuesta.pdf_zip_base64);
                        Console.WriteLine("XML EN BASE64: " + leer_respuesta.xml_zip_base64);
                        Console.WriteLine("CDR EN BASE64: " + leer_respuesta.cdr_zip_base64);
                        Console.WriteLine("CODIGO QR: " + leer_respuesta.cadena_para_codigo_qr);
                        Console.WriteLine("CODIGO HASH: " + leer_respuesta.codigo_hash);
                        Console.WriteLine("CODIGO DE BARRAS: " + leer_respuesta.codigo_de_barras);
                    }
                    else
                    {
                        if (leer_respuesta.errors != "Este documento ya existe en NubeFacT")
                            throw new Exception(leer_respuesta.errors);
                    }
                }
                catch (System.Exception ex)
                {
                }

                EstadoDocumento = "01";

                result = true;
            }
            catch (Exception ex)
            {
                mensaje_error = ex.Message;
                return false;
            }
            finally
            {
                GC.Collect();
            }
            return result;
        }

        public static string SendJson(string ruta, string json, string token)
        {
            try
            {
                using (var client = new WebClient())
                {
                    /// ESPECIFICAMOS EL TIPO DE DOCUMENTO EN EL ENCABEZADO
                    client.Headers[HttpRequestHeader.ContentType] = "application/json; charset=utf-8";
                    /// ASI COMO EL TOKEN UNICO
                    client.Headers[HttpRequestHeader.Authorization] = "Token token=" + token;
                    /// OBTENEMOS LA RESPUESTA
                    string respuesta = client.UploadString(ruta, "POST", json);
                    /// Y LA 'RETORNAMOS'
                    return respuesta;
                }
            }
            catch (WebException ex)
            {
                /// EN CASO EXISTA ALGUN ERROR, LO TOMAMOS
                var respuesta = new StreamReader(ex.Response.GetResponseStream()).ReadToEnd();
                /// Y LO 'RETORNAMOS'
                return respuesta;
            }
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

                            #region CONSULTAR COMPROBANTE ESTELA
                            Entities.Estela.Documento oBeFEDOC = new Entities.Estela.Documento();
                            oBeFEDOC.DocEntry = document.DocEntry.ToString();
                            SAPDI.ObtenerDatosCPEEstela(ref oBeFEDOC, document.Tabla);
                            Consulta consulta = new Consulta();
                            consulta.operacion = "consultar_comprobante";
                            consulta.tipo_de_comprobante = oBeFEDOC.tipoComprobante == "01" ? "1" : oBeFEDOC.tipoComprobante == "03" ? "2" : oBeFEDOC.tipoComprobante == "07" ? "3" : "4";
                            consulta.serie = "FFF1"; // oBeFEDOC.serie;
                            consulta.numero = oBeFEDOC.correlativo;

                            try
                            {
                                string json = JsonConvert.SerializeObject(consulta, Formatting.Indented);
                                Console.WriteLine(json);
                                byte[] bytes = Encoding.Default.GetBytes(json);
                                string json_en_utf_8 = Encoding.UTF8.GetString(bytes);

                                string json_de_respuesta = SendJson(Globals.URL_OSE, json_en_utf_8, Globals.TOKEN);
                                dynamic r = JsonConvert.DeserializeObject<Respuesta>(json_de_respuesta);
                                string r2 = JsonConvert.SerializeObject(r, Formatting.Indented);
                                dynamic json_r_in = JsonConvert.DeserializeObject<Respuesta>(r2);

                                dynamic respuesta = JsonConvert.DeserializeObject<Respuesta>(json_de_respuesta);
                                string mensaje = string.Empty;

                                if (respuesta.errors == null)
                                {
                                    if (respuesta.aceptada_por_sunat)
                                    {
                                        nombreArchivo = oBeFEDOC.emisor.ruc + "_" + oBeFEDOC.tipoComprobante + "_" + oBeFEDOC.serie + "_" + oBeFEDOC.correlativo;
                                        respuestaanexo = SAPDI.AnexarDocumentosCpe(document.DocEntry, nombreArchivo, respuesta.enlace_del_pdf, respuesta.enlace_del_xml, respuesta.enlace_del_cdr, document.ObjType, Globals.oCompany, ref mensajeerror);
                                        SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, respuesta.sunat_description, "AC_03", Globals.oCompany);
                                    }

                                }
                                else
                                {
                                    if (respuesta.errors != "Este documento ya existe en NubeFacT")
                                    {
                                        mensaje = "Consulta Estado: " + respuesta.errors;
                                        SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, mensaje, "01", Globals.oCompany);
                                    }
                                }

                            }
                            catch (System.Exception ex)
                            {
                            }
                            #endregion
                        }
                        break;
                    case "20":
                        foreach (Comprobante document in listaComp)
                        {
                            string mensajeerror = "";
                            string nombreArchivo = "";
                            string pdf_url = "";
                            string xml_url = "";
                            string cdr_url = "";
                            bool respuestaanexo = false;

                            #region CONSULTAR COMPROBANTE ESTELA
                            Entities.Estela.Retencion oDocumentoRetFE1 = new Retencion();
                            SAPDI.ObtenerDatosCRE(document.DocEntry, ref oDocumentoRetFE1);
                            DateTime date1 = DateTime.Now;


                            try
                            {

                            }
                            catch (System.Exception ex)
                            {
                            }
                            #endregion
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
            }
        }
    }
}
