using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using ServicioConectorFE.Core;
using ServicioConectorFE.Entities.Estela;
using ServicioConectorFE.Entities.SAP;
using ServicioConectorFE.Framework;
using ServicioConectorFE.IntegradorEstelaAdjunto;
using ServicioConectorFE.IntegradorEstelaBaja;

namespace ServicioConectorFE.OSE
{
    public class Estela
    {
        public static bool RegistrarRetencion(string docmun, int objectype, string tipodocumento, string Tabla, int docentry, ref string EstadoDocumento, ref string Folio, ref string mensaje_error)
        {
            bool result = false;

            Entities.Estela.DocumentoRetencionFE oDocumentoRetFE = new Entities.Estela.DocumentoRetencionFE();

            try
            {
                DateTime date1 = DateTime.Now;
                SAPDI.ObtenerDatosCRE(docentry, ref oDocumentoRetFE);
                String Cadena = string.Empty;

                Globals.ValidaCamposCRE(oDocumentoRetFE);

                switch (Globals.ProveedorOSE)
                {
                    case "3": //Estela
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
                            if (string.IsNullOrEmpty(oDocumentoRetFE.correoAdquiriente))
                            {

                            }
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
                        break;
                }
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
                        oDetalle.IgvTipo = detalle.Field<string>("U_SMC_TIPOAFECT_FE").ToString();
                        oDetalle.descuentoPct = Globals.RetirarMonedayDecimal(detalle.Field<string>("DiscPrcnt"));
                        oDetalle.OperacionAfecta = oDetalle.CodigoImpuesto == "IGV" ? "A" : oDetalle.CodigoImpuesto == "INA" ? "I" : oDetalle.CodigoImpuesto == "EXO" ? "E" : ""; // detalle.Field<string>("U_BPP_OPER").ToString();
                        oDetalle.TaxOnly = detalle.Field<string>("TaxOnly").ToString();
                        oDetalle.Glosa = detalle.Field<string>("U_EXA_GLOSA").ToString();

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
                        oDetalle.IgvTipo = detalle.Field<string>("U_SMC_TIPOAFECT_FE").ToString();
                        oDetalle.descuentoPct = Globals.RetirarMonedayDecimal(detalle.Field<string>("DiscPrcnt"));
                        oDetalle.OperacionAfecta = oDetalle.CodigoImpuesto == "IGV" ? "A" : oDetalle.CodigoImpuesto == "INA" ? "I" : oDetalle.CodigoImpuesto == "EXO" ? "E" : ""; // detalle.Field<string>("U_BPP_OPER").ToString();
                        oDetalle.TaxOnly = detalle.Field<string>("TaxOnly").ToString();
                        oDetalle.Glosa = detalle.Field<string>("U_EXA_GLOSA").ToString();

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
                        Globals.ObtenerAnticiposEstela(ref oBeFEDOC, Tabla, Convert.ToInt32(oBeFEDOC.DocEntry));
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

                oBeFEDOC.valorIgv = Convert.ToDouble(oBeFEDOC.valorIgv) < 0 ? "0.00" : oBeFEDOC.valorIgv;

                //Encabezado
                Cadena += "A;CODI_EMPR;;" + Globals.IdEmpresa + Environment.NewLine;
                Cadena += "A;TipoDTE;;" + oBeFEDOC.tipoComprobante + Environment.NewLine;
                Cadena += "A;Serie;;" + oBeFEDOC.serie + Environment.NewLine;
                Cadena += "A;Correlativo;;" + oBeFEDOC.correlativo + Environment.NewLine;
                Cadena += "A;FchEmis;;" + oBeFEDOC.fechaEmision + Environment.NewLine;
                Cadena += "A;HoraEmision;;" + oBeFEDOC.horaEmision + Environment.NewLine;
                Cadena += "A;TipoMoneda;;" + oBeFEDOC.tipoMoneda + Environment.NewLine;

                //Datos del emisor
                Cadena += "A;RUTEmis;;" + oBeFEDOC.emisor.ruc + Environment.NewLine;
                Cadena += "A;TipoRucEmis;;" + oBeFEDOC.emisor.tipoIdentificacion + Environment.NewLine;
                Cadena += "A;NomComer;;" + oBeFEDOC.emisor.nombreComercial + Environment.NewLine;
                Cadena += "A;RznSocEmis;;" + oBeFEDOC.emisor.razonSocial + Environment.NewLine;
                Cadena += "A;ComuEmis;;" + oBeFEDOC.emisor.ubigeo + Environment.NewLine;
                Cadena += "A;DirEmis;;" + oBeFEDOC.emisor.direccion.Replace(Environment.NewLine, "") + Environment.NewLine;
                Cadena += "A;ProviEmis;;" + oBeFEDOC.emisor.provincia + Environment.NewLine;
                Cadena += "A;DeparEmis;;" + oBeFEDOC.emisor.departamento + Environment.NewLine;
                Cadena += "A;DistriEmis;;" + oBeFEDOC.emisor.distrito + Environment.NewLine;
                Cadena += "A;CodigoLocalAnexo;;0000" + Environment.NewLine;

                //Datos del receptor
                Cadena += "A;TipoRutReceptor;;" + oBeFEDOC.adquiriente.tipoIdentificacion + Environment.NewLine;
                Cadena += "A;RUTRecep;;" + oBeFEDOC.adquiriente.numeroIdentificacion + Environment.NewLine;
                Cadena += "A;RznSocRecep;;" + oBeFEDOC.adquiriente.nombre + Environment.NewLine;
                Cadena += "A;DirRecepUbiGeo;;" + oBeFEDOC.adquiriente.ubigeo + Environment.NewLine;
                Cadena += "A;DirRecep;;" + oBeFEDOC.adquiriente.direccionFiscal.Replace(Environment.NewLine, "") + Environment.NewLine;
                //Cadena += "A;DirRecepProvincia;;" + oBeFEDOC.adquiriente. + Environment.NewLine;
                //Cadena += "A;DirRecepDepartamento;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                //Cadena += "A;DirRecepDistrito;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                //Cadena += "A;DirRecepCodPais;;" + oBeFEDOC.DocEntry + Environment.NewLine;

                //Datos de NC
                if (oBeFEDOC.tipoComprobante == "07" || oBeFEDOC.tipoComprobante == "08")
                {
                    Cadena += "A;CodigoAutorizacion;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                    Cadena += "A;Sustento;;" + oBeFEDOC.notaSustento + Environment.NewLine;
                    Cadena += "A;TipoNotaCredito;;" + oBeFEDOC.notaTipo + Environment.NewLine;
                }

                //Totales
                Cadena += "A;MntNeto;;" + (MntGrabado + MntInafecto + MntExonerado) + Environment.NewLine;
                Cadena += "A;MntExe;;" + MntInafecto + Environment.NewLine;
                Cadena += "A;MntExo;;" + MntExonerado + Environment.NewLine;
                Cadena += "A;MntTotalIgv;;" + MntIGV + Environment.NewLine;
                Cadena += "A;MntTotGrat;;" + MntTransGratuita + Environment.NewLine;
                //Cadena += "A;MntTotBoni;;" + MntTransGratuita + Environment.NewLine;
                Cadena += "A;MntTotAnticipo;;" + oBeFEDOC.valorTotalAnticipos + Environment.NewLine;
                Cadena += "A;MntTotal;;" + oBeFEDOC.importeTotal + Environment.NewLine;
                Cadena += "A;TotDscNA ;;" + oBeFEDOC.valorDescuentoGlobal + Environment.NewLine;
                Cadena += "A;TotCrgNA;;" + oBeFEDOC.otrosCargos + Environment.NewLine;
                Cadena += "A;MntRedondeo;;" + oBeFEDOC.redondeo + Environment.NewLine;


                //Leyendas
                //Cadena += "A;1000;;" + Globals.MontoEnLetras(oBeFEDOC.importeTotal, Globals.ObtenerDescripcionMoneda(oBeFEDOC.tipoMoneda)) + Environment.NewLine;

                //Retención
                /*Cadena += "A;CodRetencion;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;MntImpRetencion;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;ObsRetencion;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;MntRetencion;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;MntTotalMenosRetencion;;" + oBeFEDOC.DocEntry + Environment.NewLine;

                //Percepción
                Cadena += "A;IndPercepcion;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;CodPercepcion;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;MntImpPercepcion;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;ObsPercepcion;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;MntPercepcion;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;MntTotalMasPercepcion;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;IndEmiExcepcional;;" + oBeFEDOC.DocEntry + Environment.NewLine;*/

                //Factura Guía


                //Otros Conceptos Sunat
                if (oBeFEDOC.tipoComprobante == "01" || oBeFEDOC.tipoComprobante == "03")
                {
                    Cadena += "A;TipoOperacion;;" + oBeFEDOC.tipoOperacion + Environment.NewLine;
                    Cadena += "A;FechVencFact;;" + oBeFEDOC.fechaVencimiento + Environment.NewLine;
                }
                //Direccion de entrega
                /*Cadena += "A;DirEntregaUbiGeo;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;DirEntrega;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;DirEntregaUrbaniza;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;DirEntregaProvincia;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;DirEntregaDepartamento;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;DirEntregaDistrito;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                Cadena += "A;DirEntregaCodPais;;" + oBeFEDOC.DocEntry + Environment.NewLine;*/

                //Impuestos/Retenciones Globales
                if (Convert.ToDouble(oBeFEDOC.valorOperacionesGravadas) > 0)
                {
                    Cadena += "A2;CodigoImpuesto;;1000" + Environment.NewLine;
                    Cadena += "A2;MontoImpuesto;;" + oBeFEDOC.Detalles.Sum(x => Convert.ToDouble(x.IgvMonto)) + Environment.NewLine;
                    Cadena += "A2;TasaImpuesto;;" + oBeFEDOC.Detalles[0].tasaIGV + Environment.NewLine;
                    Cadena += "A2;MontoImpuestoBase;;" + oBeFEDOC.valorOperacionesGravadas + Environment.NewLine;
                }
                if (Convert.ToDouble(oBeFEDOC.valorOperacionesInafectas) > 0)
                {
                    Cadena += "A2;CodigoImpuesto;;9998" + Environment.NewLine;
                    Cadena += "A2;MontoImpuesto;;0.00" + Environment.NewLine;
                    Cadena += "A2;TasaImpuesto;;0" + Environment.NewLine;
                    Cadena += "A2;MontoImpuestoBase;;" + oBeFEDOC.valorOperacionesInafectas + Environment.NewLine;
                }
                if (Convert.ToDouble(oBeFEDOC.valorOperacionesExoneradas) > 0)
                {
                    Cadena += "A2;CodigoImpuesto;;9997" + Environment.NewLine;
                    Cadena += "A2;MontoImpuesto;;0.00" + Environment.NewLine;
                    Cadena += "A2;TasaImpuesto;;0" + Environment.NewLine;
                    Cadena += "A2;MontoImpuestoBase;;" + oBeFEDOC.valorOperacionesExoneradas + Environment.NewLine;
                }
                if (Convert.ToDouble(oBeFEDOC.valorOperacionesGratuitas) > 0)
                {
                    Cadena += "A2;CodigoImpuesto;;9996" + Environment.NewLine;
                    Cadena += "A2;MontoImpuesto;;" + oBeFEDOC.Detalles.Sum(x => Convert.ToDouble(x.IgvMontoGratuito)) + Environment.NewLine;
                    Cadena += "A2;TasaImpuesto;;" + oBeFEDOC.tasaIgv + Environment.NewLine;
                    Cadena += "A2;MontoImpuestoBase;;" + oBeFEDOC.valorOperacionesGratuitas + Environment.NewLine;
                }

                //Detalle Detracciones y otros
                if (oBeFEDOC.tipoComprobante == "01" || oBeFEDOC.tipoComprobante == "03" || oBeFEDOC.tipoComprobante == "08")
                {
                    if (oBeFEDOC.TieneDetraccion == "Y")
                    {
                        Cadena += "A3;codiDetraccion;1;2003" + Environment.NewLine;
                        Cadena += "A3;ValorDetraccion;1;0" + Environment.NewLine;
                        Cadena += "A3;MntDetraccion;1;" + oBeFEDOC.totalDetraccion + Environment.NewLine;
                        Cadena += "A3;PorcentajeDetraccion;1;" + oBeFEDOC.porcentajeDetraccion + Environment.NewLine;

                        Cadena += "A3;codiDetraccion;2;3000" + Environment.NewLine;
                        Cadena += "A3;ValorDetraccion;2;" + oBeFEDOC.codigoDetraccion + Environment.NewLine;
                        Cadena += "A3;MntDetraccion;2;0" + Environment.NewLine;
                        Cadena += "A3;PorcentajeDetraccion;2;0" + Environment.NewLine;

                        Cadena += "A3;codiDetraccion;3;3001" + Environment.NewLine;
                        Cadena += "A3;ValorDetraccion;3;" + (Globals.IdEmpresa == "1" ? "0000-5017939" : "0000-760994") + Environment.NewLine;
                        Cadena += "A3;MntDetraccion;3;0" + Environment.NewLine;
                        Cadena += "A3;PorcentajeDetraccion;3;0" + Environment.NewLine;
                    }

                    string TotalCuota = oBeFEDOC.TieneDetraccion == "Y" ? Globals.SumarCadenas(oBeFEDOC.importeTotal, "-" + (oBeFEDOC.tipoMoneda == "SOL" || oBeFEDOC.tipoMoneda == "PEN" ? oBeFEDOC.totalDetraccion : (Math.Round(Convert.ToDouble(oBeFEDOC.totalDetraccion) / Convert.ToDouble(oBeFEDOC.TipoCambio)), 2).ToString())) : oBeFEDOC.importeTotal;

                    //Información de Forma de pago
                    Cadena += "A;FormaPago;;" + (oBeFEDOC.credito == "Y" ? "Credito" : "Contado") + Environment.NewLine;
                    Cadena += "A;MontoNetoPendPago;;" + (oBeFEDOC.credito == "Y" ? TotalCuota : "0.00") + Environment.NewLine;

                    if (oBeFEDOC.credito == "Y")
                    {
                        Cadena += "A5;Cuota;1;1" + Environment.NewLine;
                        Cadena += "A5;MontoCuota;1;" + TotalCuota + Environment.NewLine;
                        Cadena += "A5;FechaVencCuota;1;" + oBeFEDOC.fechaVencimiento + Environment.NewLine;
                    }
                }



                for (int i = 0; i < oBeFEDOC.Detalles.Count; i++)
                {
                    Cadena += "B;NroLinDet;" + (i + 1) + ";" + (i + 1) + Environment.NewLine;
                    Cadena += "B;VlrCodigo;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    Cadena += "B;QtyItem;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].cantidadUnidades + Environment.NewLine;
                    Cadena += "B;UnmdItem;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].unidadMedida + Environment.NewLine;
                    Cadena += "B;NmbItem;" + (i + 1) + ";" + (oBeFEDOC.Detalles[i].descripcion + " " + oBeFEDOC.Detalles[i].Glosa).Trim() + Environment.NewLine;
                    Cadena += "B;PrcItem;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].precioUnitario + Environment.NewLine;
                    Cadena += "B;PrcItemSinIgv;" + (i + 1) + ";" + (oBeFEDOC.Detalles[i].TaxOnly == "N" ? oBeFEDOC.Detalles[i].valorUnitario : "0.00") + Environment.NewLine;
                    Cadena += "B;MontoItem;" + (i + 1) + ";" + (oBeFEDOC.Detalles[i].TaxOnly == "N" ? oBeFEDOC.Detalles[i].valorItem : "0.00") + Environment.NewLine;
                    Cadena += "B;DescuentoMonto;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].descuento + Environment.NewLine;
                    Cadena += "B;IndExe;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].IgvTipo + Environment.NewLine;
                    Cadena += "B;CodigoTipoIgv;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].CodigoTributo + Environment.NewLine;
                    Cadena += "B;TasaIgv;" + (i + 1) + ";" + (oBeFEDOC.Detalles[i].TaxOnly == "N" ? oBeFEDOC.Detalles[i].tasaIGV : 0.00) + Environment.NewLine;
                    Cadena += "B;ImpuestoIgv;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].IgvMonto + Environment.NewLine;
                    //Cadena += "B;CodigoIsc;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;CodigoTipoIsc;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;MontoIsc;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;TasaIsc;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    Cadena += "B;MontoBaseImp;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].valorItem + Environment.NewLine;
                    //Cadena += "B;MontoBaseIsc;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    Cadena += "B;CodigoProductoSunat;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;CodigoProductoGS1;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;GS1TipoGTIN;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;CodigoOtrosTributos;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;MontoOtrosTributos;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;MontoBaseOtrosTributos;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;TasaOtrosTributos;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;CodigoTributoBolsaPlastica;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;MontoTributoBolsaPlastica;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;MontoUnitarioBolsaPlastica;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    //Cadena += "B;CantidadBolsaPlastica;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;

                    //Cargos y Descuentos
                    /*Cadena += "B1;NroLinDet;" + (i + 1) + ";" + (i + 1) + Environment.NewLine;
                    Cadena += "B1;DescCargoDescuento;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    Cadena += "B1;IndCargoDescuento;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    Cadena += "B1;CodigoCargoDescuento;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    Cadena += "B1;FactorCargoDescuento;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    Cadena += "B1;MontoCargoDescuento;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    Cadena += "B1;MBaseCargoDescuento;" + (i + 1) + ";" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;*/

                    //Propiedades adicionales por Item
                    /*Cadena += "B2;NroLinDet;;" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    Cadena += "B2;CodConTrib;;" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    Cadena += "B2;NomConTrib;;" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    Cadena += "B2;ValConTrib;;" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;
                    Cadena += "B2;FechaConTrib;;" + oBeFEDOC.Detalles[i].codigo + Environment.NewLine;*/

                }

                //Descuentos y recargos globales
                if (oBeFEDOC.valorDescuentoGlobal != "0.00")
                {
                    Cadena += "C;NroLinDR;;1" + Environment.NewLine;
                    Cadena += "C;TpoMov;;D" + Environment.NewLine;
                    Cadena += "C;ValorDR;;" + oBeFEDOC.valorDescuentoGlobal + Environment.NewLine;
                    Cadena += "C;IndCargoDescuento;;0" + Environment.NewLine;
                    Cadena += "C;CodigoCargoDescuento;;02" + Environment.NewLine;
                    //Cadena += "C;FactorCargoDescuento;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                    Cadena += "C;MontoCargoDescuento;;" + oBeFEDOC.valorDescuentoGlobal + Environment.NewLine;
                    Cadena += "C;MBaseCargoDescuento;;" + oBeFEDOC.valorDescuentoGlobal + Environment.NewLine;
                }

                if (oBeFEDOC.otrosCargos != "0.00")
                {
                    Cadena += "C;NroLinDR;;1" + Environment.NewLine;
                    Cadena += "C;TpoMov;;R" + Environment.NewLine;
                    Cadena += "C;ValorDR;;" + oBeFEDOC.otrosCargos + Environment.NewLine;
                    Cadena += "C;IndCargoDescuento;;1" + Environment.NewLine;
                    Cadena += "C;CodigoCargoDescuento;;50" + Environment.NewLine;
                    //Cadena += "C;FactorCargoDescuento;;" + oBeFEDOC.DocEntry + Environment.NewLine;
                    Cadena += "C;MontoCargoDescuento;;" + oBeFEDOC.otrosCargos + Environment.NewLine;
                    Cadena += "C;MBaseCargoDescuento;;" + oBeFEDOC.otrosCargos + Environment.NewLine;
                }

                //Referencias
                if (oBeFEDOC.tipoComprobante == "07" || oBeFEDOC.tipoComprobante == "08")
                {
                    if (oBeFEDOC.ListReferencia.Count == 0)
                        throw new Exception("Debe llenar los datos del documento referenciado");
                    if (string.IsNullOrEmpty(oBeFEDOC.ListReferencia[0].tipoDocumento))
                        throw new Exception("Debe llenar el tipo del documento referenciado");
                    if (string.IsNullOrEmpty(oBeFEDOC.ListReferencia[0].serie))
                        throw new Exception("Debe llenar la serie del documento referenciado");
                    if (string.IsNullOrEmpty(oBeFEDOC.ListReferencia[0].correlativo))
                        throw new Exception("Debe llenar el correlativo del documento referenciado");

                    Cadena += "D;NroLinRef;1;1" + Environment.NewLine;
                    Cadena += "D;TpoDocRef;1;" + oBeFEDOC.ListReferencia[0].tipoDocumento + Environment.NewLine;
                    Cadena += "D;SerieRef;1;" + oBeFEDOC.ListReferencia[0].serie + Environment.NewLine;
                    Cadena += "D;FolioRef;1;" + oBeFEDOC.ListReferencia[0].correlativo + Environment.NewLine;

                    if (!string.IsNullOrEmpty(oBeFEDOC.ListReferencia[0].fechaEmision))
                    {
                        Cadena += "E;TipoAdicSunat;4;01" + Environment.NewLine;
                        Cadena += "E;NmrLineasAdicSunat;4;04" + Environment.NewLine;
                        Cadena += "E;DescripcionAdicsunat;4;" + Convert.ToDateTime(oBeFEDOC.ListReferencia[0].fechaEmision).ToString("dd/MM/yyyy") + Environment.NewLine;
                    }

                    if (!string.IsNullOrEmpty(oBeFEDOC.fechaVencimiento))
                    {
                        Cadena += "E;TipoAdicSunat;5;01" + Environment.NewLine;
                        Cadena += "E;NmrLineasAdicSunat;5;05" + Environment.NewLine;
                        Cadena += "E;DescripcionAdicsunat;5;" + Convert.ToDateTime(oBeFEDOC.fechaVencimiento).ToString("dd/MM/yyyy") + Environment.NewLine;
                    }
                }

                //Seccion personalizados
                int index = 1;
                if (!string.IsNullOrEmpty(oBeFEDOC.Responsable))
                {
                    Cadena += "E;TipoAdicSunat;" + index + ";01" + Environment.NewLine;
                    Cadena += "E;NmrLineasAdicSunat;" + index + ";0" + index + Environment.NewLine;
                    Cadena += "E;DescripcionAdicsunat;" + index + ";" + oBeFEDOC.Responsable + Environment.NewLine;
                }
                index++;
                if (!string.IsNullOrEmpty(oBeFEDOC.condicionPagoDesc))
                {
                    Cadena += "E;TipoAdicSunat;" + index + ";01" + Environment.NewLine;
                    Cadena += "E;NmrLineasAdicSunat;" + index + ";0" + index + Environment.NewLine;
                    Cadena += "E;DescripcionAdicsunat;" + index + ";" + oBeFEDOC.condicionPagoDesc + Environment.NewLine;
                }
                index++;
                if (!string.IsNullOrEmpty(oBeFEDOC.Comments))
                {
                    Cadena += "E;TipoAdicSunat;" + index + ";01" + Environment.NewLine;
                    Cadena += "E;NmrLineasAdicSunat;" + index + ";0" + index + Environment.NewLine;
                    Cadena += "E;DescripcionAdicsunat;" + index + ";" + oBeFEDOC.Comments + Environment.NewLine;
                }

                //Seccion de correos M
                if (!string.IsNullOrEmpty(oBeFEDOC.adquiriente.email))
                {
                    var correos = oBeFEDOC.adquiriente.email.Split(';').ToList();
                    for (int i = 0; i < correos.Count; i++)
                    {
                        Cadena += "M;NroLinMail;;" + (i + 1) + Environment.NewLine;
                        Cadena += "M;MailEnvi;;" + correos[i].Trim() + Environment.NewLine;
                    }
                }

                try
                {
                    System.IO.File.WriteAllText(Globals.PathArchivos + oBeFEDOC.emisor.ruc + "_" + oBeFEDOC.tipoComprobante + "_" + oBeFEDOC.serie + "_" + oBeFEDOC.correlativo + ".TXT", Cadena, System.Text.Encoding.Default);
                    Folio = oBeFEDOC.serie + "-" + oBeFEDOC.correlativo;
                    //SAPDI.ActualizarEstadoFE(objectype, docentry, "Documento enviado a OSE", "01", Globals.oCompany);
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
                //SAPDI.ActualizarEstadoFE(objectype, docentry, ex.Message, "", Globals.oCompany);
                return false;
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

                            #region CONSULTAR COMPROBANTE ESTELA
                            Entities.Estela.Documento oBeFEDOC = new Entities.Estela.Documento();
                            System.Data.DataTable dtDetalles = new System.Data.DataTable();
                            DateTime date1 = DateTime.Now;

                            oBeFEDOC.DocEntry = document.DocEntry.ToString();

                            SAPDI.ObtenerDatosCPEEstela(ref oBeFEDOC, document.Tabla);
                            oBeFEDOC.Detalles = new List<Entities.Estela.DocLine>();
                            SAPDI.ObtenerDetalleDocumento(ref dtDetalles, document.Tabla, document.DocEntry);

                            try
                            {
                                nombreArchivo = oBeFEDOC.emisor.ruc + "_" + oBeFEDOC.tipoComprobante + "_" + oBeFEDOC.serie + "_" + oBeFEDOC.correlativo;

                                var clientEstela = new IntegradorEstelaEstado.Service1SoapClient();
                                var response3 = clientEstela.ConsultaEstado(oBeFEDOC.emisor.ruc, Convert.ToInt32(oBeFEDOC.tipoComprobante), Convert.ToInt32(oBeFEDOC.correlativo), oBeFEDOC.serie);

                                List<string> rpta = response3.ToString().Split(';').ToList();

                                clientEstela.Close();

                                switch (rpta[0])
                                {
                                    case "DOK":
                                        var bindingEstela = new BasicHttpBinding(BasicHttpSecurityMode.Transport)
                                        {
                                            MaxReceivedMessageSize = 10485760,
                                            ReaderQuotas = { MaxArrayLength = 10485760 }
                                        };

                                        var endpointEstela = new EndpointAddress("https://pruebas.ecomprobantes.pe/wssConsultaDocPeru/consultaDocumentosPeru.asmx");
                                        var factoryEstela = new ChannelFactory<consultaDocumentosPeruSoap>(bindingEstela, endpointEstela);
                                        var clienteEstela2 = factoryEstela.CreateChannel();

                                        var requestBodyPDF = new getDocumentoPDFRequestBody
                                        (
                                            _ruttEmpr: oBeFEDOC.emisor.ruc,
                                            _folioDTE: oBeFEDOC.correlativo,
                                            _tipoDTE: oBeFEDOC.tipoComprobante,
                                            _serieInte: oBeFEDOC.serie,
                                            _fechaDTE: oBeFEDOC.fechaEmision,
                                            _monTotal: oBeFEDOC.importeTotal
                                        );
                                        var requestPDF = new getDocumentoPDFRequest(requestBodyPDF);
                                        var responsePDF = clienteEstela2.getDocumentoPDF(requestPDF);
                                        if (string.IsNullOrEmpty(responsePDF.Body.getDocumentoPDFResult[0])) cdr_url = string.Empty;
                                        else pdf_url = responsePDF.Body.getDocumentoPDFResult[1];

                                        var requestBodyXML = new getDocumentoXMLRequestBody
                                        (
                                           _ruttEmpr: oBeFEDOC.emisor.ruc,
                                            _folioDTE: oBeFEDOC.correlativo,
                                            _tipoDTE: oBeFEDOC.tipoComprobante,
                                            _serieInte: oBeFEDOC.serie,
                                            _fechaDTE: oBeFEDOC.fechaEmision,
                                            _monTotal: oBeFEDOC.importeTotal
                                        );
                                        var requestXML = new getDocumentoXMLRequest(requestBodyXML);
                                        var responseXML = clienteEstela2.getDocumentoXML(requestXML);
                                        if (string.IsNullOrEmpty(responseXML.Body.getDocumentoXMLResult[0])) cdr_url = string.Empty;
                                        else xml_url = responseXML.Body.getDocumentoXMLResult[1];

                                        var requestBodyCRD = new getDocumentoCDRRequestBody
                                        (
                                           _ruttEmpr: oBeFEDOC.emisor.ruc,
                                            _folioDTE: oBeFEDOC.correlativo,
                                            _tipoDTE: oBeFEDOC.tipoComprobante,
                                            _serieInte: oBeFEDOC.serie,
                                            _fechaDTE: oBeFEDOC.fechaEmision,
                                            _monTotal: oBeFEDOC.importeTotal
                                        );
                                        var requestCRD = new getDocumentoCDRRequest(requestBodyCRD);
                                        var responseCRD = clienteEstela2.getDocumentoCDR(requestCRD);
                                        if (string.IsNullOrEmpty(responseCRD.Body.getDocumentoCDRResult[0])) cdr_url = string.Empty;
                                        else cdr_url = responseCRD.Body.getDocumentoCDRResult[1];

                                        ((IClientChannel)clienteEstela2).Close();

                                        #region Anexar Datos Comprobantes

                                        respuestaanexo = SAPDI.AnexarDocumentosCpe(document.DocEntry, nombreArchivo, pdf_url, xml_url, cdr_url, document.ObjType, Globals.oCompany, ref mensajeerror);
                                        SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, rpta[1], "AC_03", Globals.oCompany);
                                        #endregion
                                        break;
                                    case "ING":
                                        throw new System.Exception(rpta[4]);
                                    case "ERR":
                                    case "RCH":
                                        //Campo mensaje crear como texto
                                        string mensaje = rpta[1] + ":" + rpta[2];
                                        SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, mensaje, "RC_05", Globals.oCompany);
                                        break;
                                    case "ANUL":
                                        throw new System.Exception("El documento en consulta se encuentra anulado");
                                    default:
                                        throw new System.Exception(rpta[1] + ": " + rpta[2]);
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
                            DocumentoRetencionFE oDocumentoRetFE1 = new DocumentoRetencionFE();
                            SAPDI.ObtenerDatosCRE(document.DocEntry, ref oDocumentoRetFE1);
                            DateTime date1 = DateTime.Now;


                            try
                            {
                                nombreArchivo = oDocumentoRetFE1.numeroDocumentoEmisor + "_" + oDocumentoRetFE1.tipoDocumento + "_" + oDocumentoRetFE1.serieNumeroRetencion.Replace('-', '_');

                                var clientEstela = new IntegradorEstelaEstado.Service1SoapClient();
                                var response3 = clientEstela.ConsultaEstado(oDocumentoRetFE1.numeroDocumentoEmisor, Convert.ToInt32(oDocumentoRetFE1.tipoDocumento), Convert.ToInt32(oDocumentoRetFE1.serieNumeroRetencion.Split('-').ToList()[1]), oDocumentoRetFE1.serieNumeroRetencion.Split('-').ToList()[0]);

                                List<string> rpta = response3.ToString().Split(';').ToList();

                                clientEstela.Close();
                                string mensaje = string.Empty;
                                switch (rpta[0])
                                {
                                    case "DOK":
                                        var bindingEstela = new BasicHttpBinding(BasicHttpSecurityMode.Transport)
                                        {
                                            MaxReceivedMessageSize = 10485760,
                                            ReaderQuotas = { MaxArrayLength = 10485760 }
                                        };

                                        var endpointEstela = new EndpointAddress("https://pruebas.ecomprobantes.pe/wssConsultaDocPeru/consultaDocumentosPeru.asmx");
                                        var factoryEstela = new ChannelFactory<consultaDocumentosPeruSoap>(bindingEstela, endpointEstela);
                                        var clienteEstela2 = factoryEstela.CreateChannel();

                                        var requestBodyPDF = new getDocumentoPDFRequestBody
                                        (
                                            _ruttEmpr: oDocumentoRetFE1.numeroDocumentoEmisor,
                                            _folioDTE: oDocumentoRetFE1.serieNumeroRetencion.Split('-').ToList()[1],
                                            _tipoDTE: oDocumentoRetFE1.tipoDocumento,
                                            _serieInte: oDocumentoRetFE1.serieNumeroRetencion.Split('-').ToList()[0],
                                            _fechaDTE: oDocumentoRetFE1.fechaEmision,
                                            _monTotal: oDocumentoRetFE1.importeTotalRetenido
                                        );
                                        var requestPDF = new getDocumentoPDFRequest(requestBodyPDF);
                                        var responsePDF = clienteEstela2.getDocumentoPDF(requestPDF);
                                        if (string.IsNullOrEmpty(responsePDF.Body.getDocumentoPDFResult[0])) cdr_url = string.Empty;
                                        else pdf_url = responsePDF.Body.getDocumentoPDFResult[1];

                                        var requestBodyXML = new getDocumentoXMLRequestBody
                                        (
                                            _ruttEmpr: oDocumentoRetFE1.numeroDocumentoEmisor,
                                            _folioDTE: oDocumentoRetFE1.serieNumeroRetencion.Split('-').ToList()[1],
                                            _tipoDTE: oDocumentoRetFE1.tipoDocumento,
                                            _serieInte: oDocumentoRetFE1.serieNumeroRetencion.Split('-').ToList()[0],
                                            _fechaDTE: oDocumentoRetFE1.fechaEmision,
                                            _monTotal: oDocumentoRetFE1.importeTotalRetenido
                                        );
                                        var requestXML = new getDocumentoXMLRequest(requestBodyXML);
                                        var responseXML = clienteEstela2.getDocumentoXML(requestXML);
                                        if (string.IsNullOrEmpty(responseXML.Body.getDocumentoXMLResult[0])) cdr_url = string.Empty;
                                        else xml_url = responseXML.Body.getDocumentoXMLResult[1];

                                        var requestBodyCRD = new getDocumentoCDRRequestBody
                                        (
                                            _ruttEmpr: oDocumentoRetFE1.numeroDocumentoEmisor,
                                            _folioDTE: oDocumentoRetFE1.serieNumeroRetencion.Split('-').ToList()[1],
                                            _tipoDTE: oDocumentoRetFE1.tipoDocumento,
                                            _serieInte: oDocumentoRetFE1.serieNumeroRetencion.Split('-').ToList()[0],
                                            _fechaDTE: oDocumentoRetFE1.fechaEmision,
                                            _monTotal: oDocumentoRetFE1.importeTotalRetenido
                                        );
                                        var requestCRD = new getDocumentoCDRRequest(requestBodyCRD);
                                        var responseCRD = clienteEstela2.getDocumentoCDR(requestCRD);
                                        if (string.IsNullOrEmpty(responseCRD.Body.getDocumentoCDRResult[0])) cdr_url = string.Empty;
                                        else cdr_url = responseCRD.Body.getDocumentoCDRResult[1];

                                        ((IClientChannel)clienteEstela2).Close();

                                        #region Anexar Datos Comprobantes

                                        respuestaanexo = SAPDI.AnexarDocumentosCpe(document.DocEntry, nombreArchivo, pdf_url, xml_url, cdr_url, document.ObjType, Globals.oCompany, ref mensajeerror);
                                        SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, rpta[1], "AC_03", Globals.oCompany);
                                        #endregion
                                        break;
                                    case "ING":
                                        throw new System.Exception(rpta[4]);
                                    case "ERR":
                                    case "NBD":
                                        mensaje = "Consulta Estado: " + rpta[1] + ":" + rpta[2];
                                        SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, mensaje, "01", Globals.oCompany);
                                        break;
                                    case "RCH":
                                        //Campo mensaje crear como texto
                                        mensaje = rpta[1] + ":" + rpta[2];
                                        SAPDI.ActualizarEstadoFE(document.ObjType, document.DocEntry, mensaje, "RC_05", Globals.oCompany);
                                        break;
                                    case "ANUL":
                                        throw new System.Exception("El documento en consulta se encuentra anulado");
                                    default:
                                        throw new System.Exception(rpta[1] + ": " + rpta[2]);
                                }
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

        public static void ValidaCamposCPE(Entities.Estela.Documento oBeFEDOC)
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
