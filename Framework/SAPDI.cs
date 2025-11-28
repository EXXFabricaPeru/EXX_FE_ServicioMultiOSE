using SAPbobsCOM;
using ServicioConectorFE.Core;
using ServicioConectorFE.Entities.Estela;
using ServicioConectorFE.Entities.SAP;
using ServicioConectorFE.Functionality;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using static System.Runtime.CompilerServices.RuntimeHelpers;
using ServicioConectorFE.OSE;
using Microsoft.Extensions.Logging.Abstractions;
using System.Security.AccessControl;
using System.Security.Permissions;
using System.Globalization;

namespace ServicioConectorFE.Framework
{
    public class SAPDI
    {
        public static string ErrMsg;
        public static int ErrCode;

        public static void Conectar(CONF oCONFIG)
        {
            try
            {
                Globals.oCompany = new SAPbobsCOM.Company();
                Globals.oCompany.Server = oCONFIG.Server;
                Globals.oCompany.CompanyDB = oCONFIG.CompanyDB;
                Globals.oCompany.UserName = oCONFIG.UserName;
                Globals.oCompany.Password = oCONFIG.Password;
                Globals.oCompany.DbServerType = (SAPbobsCOM.BoDataServerTypes)oCONFIG.DbServerType;
                Globals.oCompany.DbUserName = oCONFIG.DbUserName;
                Globals.oCompany.DbPassword = oCONFIG.DbPassword;
                Globals.oCompany.language = BoSuppLangs.ln_Spanish;
                Globals.oCompany.UseTrusted = false;

                long retVal = Globals.oCompany.Connect();
                if (retVal != 0)
                {
                    Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception(string.Format("Error de Conexión: {0}-{1}", ErrCode, ErrMsg));
                }
                Globals.IdEmpresa = oCONFIG.IdEmpresa;
                Globals.Log("Servicio conectado a la base " + oCONFIG.CompanyDB + " de SAP");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void ObtenerDatosOADM()
        {
            try
            {
                //System.Data.DataTable dt = new System.Data.DataTable();
                Globals.Query = "SELECT * FROM OADM";
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();
                if (Globals.oRec.RecordCount > 0)
                {
                    //dt = Globals.RecorsetToDataTable(Globals.oRec);
                    //Globals.oOADM.MainCurncy = string.IsNullOrEmpty(dt.Rows[0]["MainCurncy"].ToString()) ? "" : dt.Rows[0]["MainCurncy"].ToString().Trim();
                    //Globals.oOADM.SysCurrncy = string.IsNullOrEmpty(dt.Rows[0]["SysCurrncy"].ToString()) ? "" : dt.Rows[0]["SysCurrncy"].ToString().Trim();
                    //Globals.oOADM.CompnyName = string.IsNullOrEmpty(dt.Rows[0]["CompnyName"].ToString()) ? "" : dt.Rows[0]["CompnyName"].ToString().Trim().ToUpper();
                    //Globals.oOADM.AliasName = string.IsNullOrEmpty(dt.Rows[0]["AliasName"].ToString()) ? "" : dt.Rows[0]["AliasName"].ToString().Trim().ToUpper();
                    //Globals.oOADM.CompnyAddr = string.IsNullOrEmpty(dt.Rows[0]["CompnyAddr"].ToString()) ? "" : dt.Rows[0]["CompnyAddr"].ToString().Trim().ToUpper();
                    //Globals.oOADM.Country = string.IsNullOrEmpty(dt.Rows[0]["Country"].ToString()) ? "" : dt.Rows[0]["Country"].ToString().Trim().ToUpper();
                    //Globals.oOADM.TaxIdNum = string.IsNullOrEmpty(dt.Rows[0]["TaxIdNum"].ToString()) ? "" : dt.Rows[0]["TaxIdNum"].ToString().Trim().ToUpper();
                    //Globals.oOADM.RevOffice = string.IsNullOrEmpty(dt.Rows[0]["RevOffice"].ToString()) ? "" : dt.Rows[0]["RevOffice"].ToString().Trim().ToUpper();
                    //Globals.oOADM.E_Mail = string.IsNullOrEmpty(dt.Rows[0]["E_Mail"].ToString()) ? "" : dt.Rows[0]["E_Mail"].ToString().Trim().ToUpper();
                    //Globals.oOADM.Phone1 = string.IsNullOrEmpty(dt.Rows[0]["Phone1"].ToString()) ? "" : dt.Rows[0]["Phone1"].ToString().Trim().ToUpper();
                    //Globals.oOADM.Phone2 = string.IsNullOrEmpty(dt.Rows[0]["Phone2"].ToString()) ? "" : dt.Rows[0]["Phone2"].ToString().Trim().ToUpper();
                    //Globals.oOADM.Fax = string.IsNullOrEmpty(dt.Rows[0]["Fax"].ToString()) ? "" : dt.Rows[0]["Fax"].ToString().Trim().ToUpper();

                    Globals.oOADM.MainCurncy = Globals.oRec.Fields.Item("MainCurncy").Value.ToString();
                    Globals.oOADM.SysCurrncy = Globals.oRec.Fields.Item("SysCurrncy").Value.ToString();
                    Globals.oOADM.CompnyName = Globals.oRec.Fields.Item("CompnyName").Value.ToString();
                    Globals.oOADM.PrintHeader = Globals.oRec.Fields.Item("PrintHeadr").Value.ToString();
                    Globals.oOADM.AliasName = Globals.oRec.Fields.Item("AliasName").Value.ToString();
                    Globals.oOADM.CompnyAddr = Globals.oRec.Fields.Item("CompnyAddr").Value.ToString();
                    Globals.oOADM.Country = Globals.oRec.Fields.Item("Country").Value.ToString();
                    Globals.oOADM.TaxIdNum = Globals.oRec.Fields.Item("TaxIdNum").Value.ToString();
                    Globals.oOADM.RevOffice = Globals.oRec.Fields.Item("RevOffice").Value.ToString();
                    Globals.oOADM.E_Mail = Globals.oRec.Fields.Item("E_Mail").Value.ToString();
                    Globals.oOADM.Phone1 = Globals.oRec.Fields.Item("Phone1").Value.ToString();
                    Globals.oOADM.Phone2 = Globals.oRec.Fields.Item("Phone2").Value.ToString();
                    Globals.oOADM.Fax = Globals.oRec.Fields.Item("Fax").Value.ToString();
                }
                else
                {
                    throw new Exception("No se ha logrado obtener datos de configuración. Por favor verifique.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }

        public static void ObtenerDatosADM1()
        {
            try
            {
                //System.Data.DataTable dt = new System.Data.DataTable();
                Globals.Query = Properties.Resources.ObtenerADM1;
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();
                if (Globals.oRec.RecordCount > 0)
                {
                    //dt = Globals.RecorsetToDataTable(Globals.oRec);
                    //Globals.oADM1.ZipCode = dt.Rows[0]["ZipCode"].ToString().Trim();
                    //Globals.oADM1.Building = dt.Rows[0]["Building"].ToString().Trim();
                    //Globals.oADM1.County = dt.Rows[0]["County"].ToString().Trim().ToUpper();
                    //Globals.oADM1.City = dt.Rows[0]["City"].ToString().Trim().ToUpper();
                    //Globals.oADM1.State = dt.Rows[0]["State"].ToString().Trim().ToUpper();
                    //Globals.oADM1.Street = dt.Rows[0]["Street"].ToString().Trim().ToUpper();
                    //Globals.oADM1.IntrntAdrs = dt.Rows[0]["IntrntAdrs"].ToString().Trim().ToUpper();

                    Globals.oADM1.ZipCode = Globals.oRec.Fields.Item("ZipCode").Value.ToString();
                    Globals.oADM1.Building = Globals.oRec.Fields.Item("Building").Value.ToString();
                    Globals.oADM1.County = Globals.oRec.Fields.Item("County").Value.ToString();
                    Globals.oADM1.City = Globals.oRec.Fields.Item("City").Value.ToString();
                    Globals.oADM1.State = Globals.oRec.Fields.Item("State").Value.ToString();
                    Globals.oADM1.Street = Globals.oRec.Fields.Item("Street").Value.ToString();
                    Globals.oADM1.IntrntAdrs = Globals.oRec.Fields.Item("IntrntAdrs").Value.ToString();
                }
                else
                {
                    throw new Exception("No existen datos en tabla: ADM1.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }

        public static void ObtenerMonedas()
        {
            try
            {
                Globals.Query = ServicioConectorFE.Properties.Resources.ObtenerMonedas;
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();

                if (Globals.oRec.RecordCount == 0)
                    throw new Exception("No se encuentra definido códigos de monedas. Por favor verifique el maestro de monedas.");
                else
                    Globals.dtMonedas = Globals.RecorsetToDataTable(Globals.oRec);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }

        public static bool ObtenerConfig()
        {
            try
            {
                Globals.Query = "SELECT TOP 1 * FROM \"@EXX_FECONF\" ";
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();

                if (Globals.oRec.RecordCount > 0)
                {
                    Globals.URL_OSE = Globals.oRec.Fields.Item("U_EXX_WSDOCURL").Value.ToString();
                    Globals.SFTP_IP = Globals.oRec.Fields.Item("U_EXX_SFTPIP").Value.ToString();
                    //Globals.SFTP_PORT = Convert.ToInt32(Globals.oRec.Fields.Item("U_EXX_SPUERTO").Value.ToString());
                    Globals.SFTP_USER = Globals.oRec.Fields.Item("U_EXX_SUSER").Value.ToString();
                    Globals.SFTP_PASS = Globals.oRec.Fields.Item("U_EXX_SPASS").Value.ToString();
                    Globals.TOKEN = Globals.oRec.Fields.Item("U_EXX_WSTOKEN").Value.ToString();

                    Globals.ProveedorOSE = Globals.oRec.Fields.Item("U_EXX_OSE").Value.ToString();
                    Globals.PathArchivos = Globals.oRec.Fields.Item("U_EXX_PATH").Value.ToString();
                    return true;
                }
                else
                {
                    throw new Exception("No se encontró la configuración inicial para el AddOn de Facturación Electrónica.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }
        //.DocEntry, Convert.ToInt32(document.NumeroSAP), document.ObjectId, document.CodTipoDoc
        public static bool ActualizarFolioSAP(Comprobante document)
        {
            bool folioact = true;
            string FolioPref = "";
            string FolioNum = "";
            bool verfolio = true;

            try
            {
                SAPbobsCOM.Documents odoc = null; //nota de credito

                if (document.ObjectId == 203)
                {
                    if (document.CodTipoDoc == "01" || document.CodTipoDoc == "03" || document.CodTipoDoc == "07" || document.CodTipoDoc == "08")
                    {
                        Update_Folio_FactAnt(document.CodTipoDoc, Convert.ToInt32(document.NumeroSAP), document.ObjectId, document.DocEntry, ref FolioPref, ref FolioNum);
                        odoc = (SAPbobsCOM.Documents)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments);
                    }
                }
                else
                {
                    if (document.CodTipoDoc == "01" || document.CodTipoDoc == "03" || document.CodTipoDoc == "08")
                    {
                        Update_Folio_Fact(document.CodTipoDoc, Convert.ToInt32(document.NumeroSAP), document.ObjectId, document.DocEntry, ref FolioPref, ref FolioNum);
                        odoc = (SAPbobsCOM.Documents)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                    }

                    if (document.CodTipoDoc == "07")
                    {
                        Update_Folio_NotaCredito(document.CodTipoDoc, Convert.ToInt32(document.NumeroSAP), document.ObjectId, document.DocEntry, ref FolioPref, ref FolioNum);
                        odoc = (SAPbobsCOM.Documents)Globals.oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Convert.ToInt32(document.ObjectId));
                    }

                    if (document.CodTipoDoc == "20")
                    {
                        ActualizarFolioRetencion(document.DocEntry);
                        return true;
                    }

                }

                odoc.GetByKey(Convert.ToInt32(document.DocEntry));

                try
                {
                    odoc.FolioPrefixString = FolioPref;
                    odoc.FolioNumber = Convert.ToInt32(FolioNum);
                    int res = odoc.Update();

                    int temp_int = res;
                    string temp_string = "";
                    Globals.oCompany.GetLastError(out temp_int, out temp_string);

                    if (res > -1)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(odoc);
                        odoc = null;
                        GC.Collect();
                    }
                    else
                    {
                        Globals.Log("Error al asignar folio " + odoc.FolioPrefixString + "-" + odoc.FolioNumber + " al documento SAP " + document.NumeroSAP + ": " + temp_string);
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }

            return folioact;
        }

        public static void Update_Folio_FactAnt(string TipoDoc, int DocNum, int ObjectType, int DocEntry, ref string FolioPref, ref string FolioNum)
        {
            FolioNum = "0";
            FolioPref = "0";
            string nombre_Store = "EXX_Actualizar_FELFolioPrefNumFactAnt";
            SAPbobsCOM.Recordset oRecFE = default(SAPbobsCOM.Recordset);

            string Query = "";
            try
            {
                if (SAPDI.IsHana()) { Query = "CALL \"" + nombre_Store + "\" ('" + DocEntry + "','" + ObjectType + "','" + TipoDoc + "')"; }
                else { Query = "EXEC \"" + nombre_Store + "\" '" + DocEntry + "','" + ObjectType + "','" + TipoDoc + "'"; }
                oRecFE = Globals.RunQuery(Query);
                if (oRecFE != null)
                {
                    if (oRecFE.RecordCount > 0)
                    {
                        oRecFE.MoveFirst();
                        while (!oRecFE.EoF)
                        {
                            //DocEntry = oRecFE.Fields.Item("DocEntry").Value.ToString();
                            FolioPref = oRecFE.Fields.Item("FolioPref").Value.ToString();
                            FolioNum = oRecFE.Fields.Item("FolioNum").Value.ToString();
                            oRecFE.MoveNext();
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                ex.Message.ToString();
            }
            finally
            {
                oRecFE = null;
                GC.Collect();
            }
        }

        public static void Update_Folio_Fact(string TipoDoc, int DocNum, int ObjectType, int DocEntry, ref string FolioPref, ref string FolioNum)
        {
            string nombre_Store = "EXX_Actualizar_FELFolioPrefNumOINV";
            SAPbobsCOM.Recordset oRecFE = default(SAPbobsCOM.Recordset);
            string Query = "";
            try
            {
                if (SAPDI.IsHana()) { Query = "CALL \"" + nombre_Store + "\" ('" + DocEntry + "','" + ObjectType + "','" + TipoDoc + "')"; }
                else { Query = "EXEC \"" + nombre_Store + "\" '" + DocEntry + "','" + ObjectType + "','" + TipoDoc + "'"; }
                oRecFE = Globals.RunQuery(Query);
                if (oRecFE != null)
                {
                    if (oRecFE.RecordCount > 0)
                    {
                        oRecFE.MoveFirst();
                        while (!oRecFE.EoF)
                        {
                            // DocEntry = oRecFE.Fields.Item("DocEntry").Value.ToString();
                            FolioPref = oRecFE.Fields.Item("FolioPref").Value.ToString();
                            FolioNum = oRecFE.Fields.Item("FolioNum").Value.ToString();
                            oRecFE.MoveNext();
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                ex.Message.ToString();
            }
            finally
            {
                oRecFE = null;
                GC.Collect();
            }
        }

        public static void Update_Folio_NotaCredito(string TipoDoc, int DocNum, int ObjectType, int DocEntry, ref string FolioPref, ref string FolioNum)
        {

            string nombre_Store = "EXX_Actualizar_FELFolioPrefNumORIN";
            SAPbobsCOM.Recordset oRecFE = default(SAPbobsCOM.Recordset);
            string Query = "";
            try
            {
                if (SAPDI.IsHana()) { Query = "CALL \"" + nombre_Store + "\" ('" + DocEntry + "','" + ObjectType + "','" + TipoDoc + "')"; }
                else { Query = "EXEC \"" + nombre_Store + "\" '" + DocEntry + "','" + ObjectType + "','" + TipoDoc + "'"; }
                oRecFE = Globals.RunQuery(Query);
                if (oRecFE != null)
                {
                    if (oRecFE.RecordCount > 0)
                    {
                        oRecFE.MoveFirst();
                        while (!oRecFE.EoF)
                        {
                            //DocEntry = oRecFE.Fields.Item("DocEntry").Value.ToString();
                            FolioPref = oRecFE.Fields.Item("FolioPref").Value.ToString();
                            FolioNum = oRecFE.Fields.Item("FolioNum").Value.ToString();
                            oRecFE.MoveNext();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ex.Message.ToString();
            }
            finally
            {
                oRecFE = null;
                GC.Collect();
            }
        }

        public static bool ActualizarFolioRetencion(int Docentry)
        {
            string nombre_Store = "EXX_Actualizar_FolioRetencion";
            SAPbobsCOM.Recordset oRecFE = default(SAPbobsCOM.Recordset); 
            string Query = "";
            try
            {
                if (SAPDI.IsHana()) { Query = "CALL \"" + nombre_Store + "\" ('" + Docentry + "')"; }
                else { Query = "EXEC \"" + nombre_Store + "\" '" + Docentry + "'"; }
                oRecFE = Globals.RunQuery(Query);

                return true;
            }
            catch (Exception ex)
            {
                ex.Message.ToString();
                return false;
            }
            finally
            {
                oRecFE = null;
                GC.Collect();
            }
        }

        public static bool IsHana()
        {
            try
            {
                if (Globals.oCompany.DbServerType == (SAPbobsCOM.BoDataServerTypes)9)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static List<Comprobante> ListarComprobantexEmitir(DateTime FECHAINI, DateTime FECHAFIN, String TIPODOC, String ESTADO)
        {
            try
            {
                List<Comprobante> oLISTA = new List<Comprobante>();

                string query = "CALL \"EXX_COMPROBANTES_ELECTRONICOS\" ('{0}', '{1}', '{2}', '{3}')";
                query = string.Format(query, FECHAINI.ToString("yyyyMMdd"), FECHAFIN.ToString("yyyyMMdd"), TIPODOC, ESTADO);
                Globals.RunQuery(query);

                Globals.oRec.MoveFirst();
                while (!(Globals.oRec.EoF))
                {
                    Comprobante oitem = new Comprobante();

                    oitem.Tabla = Globals.oRec.Fields.Item("Tabla").Value.ToString();
                    oitem.ObjType = Convert.ToInt32(Globals.oRec.Fields.Item("ObjType").Value.ToString());
                    oitem.DocEntry = Convert.ToInt32(Globals.oRec.Fields.Item("DocEntry").Value.ToString());
                    oitem.NumeroSAP = Globals.oRec.Fields.Item("DocNum").Value.ToString();
                    oitem.Ruc = Globals.oRec.Fields.Item("Ruc").Value.ToString();
                    oitem.RazonSocial = Globals.oRec.Fields.Item("CardName").Value.ToString();
                    oitem.CodTipoDoc = Globals.oRec.Fields.Item("Indicator").Value.ToString();
                    oitem.EntryCodTipo = oitem.DocEntry + "-" + oitem.CodTipoDoc;

                    oitem.TipoDocumento = Globals.oRec.Fields.Item("TipoDocumento").Value.ToString();

                    oitem.Serie = Globals.oRec.Fields.Item("FolioPref").Value.ToString();
                    oitem.Correlativo = Globals.oRec.Fields.Item("FolioNum").Value.ToString();
                    oitem.Moneda = Globals.oRec.Fields.Item("DocCur").Value.ToString();
                    oitem.Fecha = Convert.ToDateTime(Globals.oRec.Fields.Item("DocDate").Value.ToString()).ToString("dd/MM/yyyy");
                    oitem.ObjectId = Convert.ToInt32(Globals.oRec.Fields.Item("objectId").Value.ToString());


                    oitem.Estado = Globals.oRec.Fields.Item("U_EXX_ESTADO_FE").Value.ToString();

                    oitem.Impuesto = string.Format("{0:N}", Convert.ToDecimal(Globals.oRec.Fields.Item("VatSum").Value.ToString()));
                    oitem.ImporteTotal = string.Format("{0:N}", Convert.ToDecimal(Globals.oRec.Fields.Item("DocTotal").Value.ToString()));
                    if (Globals.oRec.Fields.Item("Nota").Value.ToString().Length > 0)
                    {
                        oitem.Nota = Globals.oRec.Fields.Item("Nota").Value.ToString();
                    }
                    oitem.ReferenciaSerie = Globals.oRec.Fields.Item("ReferenciaSerie").Value.ToString();

                    oitem.Sucursal = Globals.oRec.Fields.Item("Sucursal").Value.ToString();

                    oLISTA.Add(oitem);
                    Globals.oRec.MoveNext();
                }
                return oLISTA;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static List<Baja> ListarComprobanteBaja(DateTime FECHAINI, DateTime FECHAFIN, String TIPODOC)
        {
            try
            {
                List<Baja> oLISTA = new List<Baja>();

                string query = "CALL \"EXX_BAJAS_ELECTRONICAS\" ('{0}', '{1}', '{2}')";
                query = string.Format(query, FECHAINI.ToString("yyyyMMdd"), FECHAFIN.ToString("yyyyMMdd"), TIPODOC);
                Globals.RunQuery(query);

                Globals.oRec.MoveFirst();
                while (!(Globals.oRec.EoF))
                {
                    Baja oitem = new Baja();

                    oitem.Tabla = Globals.oRec.Fields.Item("Tabla").Value.ToString();
                    oitem.ObjType = Convert.ToInt32(Globals.oRec.Fields.Item("ObjType").Value.ToString());
                    oitem.DocEntry = Convert.ToInt32(Globals.oRec.Fields.Item("DocEntry").Value.ToString());
                    oitem.ObjTypeAux = Convert.ToInt32(Globals.oRec.Fields.Item("ObjTypeAux").Value.ToString());
                    oitem.DocEntryAux = Convert.ToInt32(Globals.oRec.Fields.Item("DocEntryAux").Value.ToString());
                    oitem.Estado = Globals.oRec.Fields.Item("Estado").Value.ToString();
                    oitem.tipoComprobante = Globals.oRec.Fields.Item("tipoComprobante").Value.ToString();
                    oitem.serie = Globals.oRec.Fields.Item("serie").Value.ToString();
                    oitem.correlativo = Globals.oRec.Fields.Item("correlativo").Value.ToString();
                    oitem.fechaEmision = Convert.ToDateTime(Globals.oRec.Fields.Item("fechaEmision").Value.ToString()).ToString("yyyy-MM-dd");
                    oitem.notaSustento = Globals.oRec.Fields.Item("notaSustento").Value.ToString();

                    oLISTA.Add(oitem);
                    Globals.oRec.MoveNext();
                }
                return oLISTA;
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        public static List<Comprobante> BuscarComprobanteAnulado(DateTime FECHAINI, DateTime FECHAFIN, String TIPODOC, String ESTADO)
        {
            try
            {
                List<Comprobante> oLISTA = new List<Comprobante>();

                string query = "CALL \"EXX_COMPROBANTES_ELECTRONICOS\" ('{0}', '{1}', '{2}', '{3}')";
                query = string.Format(query, FECHAINI.ToString("yyyyMMdd"), FECHAFIN.ToString("yyyyMMdd"), TIPODOC, ESTADO);
                Globals.RunQuery(query);

                Globals.oRec.MoveFirst();
                while (!(Globals.oRec.EoF))
                {
                    Comprobante oitem = new Comprobante();

                    oitem.Tabla = Globals.oRec.Fields.Item("Tabla").Value.ToString();
                    oitem.ObjType = Convert.ToInt32(Globals.oRec.Fields.Item("ObjType").Value.ToString());
                    oitem.DocEntry = Convert.ToInt32(Globals.oRec.Fields.Item("DocEntry").Value.ToString());
                    oitem.NumeroSAP = Globals.oRec.Fields.Item("DocNum").Value.ToString();
                    oitem.Ruc = Globals.oRec.Fields.Item("Ruc").Value.ToString();
                    oitem.RazonSocial = Globals.oRec.Fields.Item("CardName").Value.ToString();
                    oitem.CodTipoDoc = Globals.oRec.Fields.Item("Indicator").Value.ToString();
                    oitem.EntryCodTipo = oitem.DocEntry + "-" + oitem.CodTipoDoc;

                    oitem.TipoDocumento = Globals.oRec.Fields.Item("TipoDocumento").Value.ToString();

                    oitem.Serie = Globals.oRec.Fields.Item("FolioPref").Value.ToString();
                    oitem.Correlativo = Globals.oRec.Fields.Item("FolioNum").Value.ToString();
                    oitem.Moneda = Globals.oRec.Fields.Item("DocCur").Value.ToString();
                    oitem.Fecha = Convert.ToDateTime(Globals.oRec.Fields.Item("TaxDate").Value.ToString()).ToString("dd/MM/yyyy");
                    oitem.ObjectId = Convert.ToInt32(Globals.oRec.Fields.Item("objectId").Value.ToString());


                    oitem.Estado = Globals.oRec.Fields.Item("U_SMC_ESTADO_FE").Value.ToString();

                    oitem.Impuesto = string.Format("{0:N}", Convert.ToDecimal(Globals.oRec.Fields.Item("VatSum").Value.ToString()));
                    oitem.ImporteTotal = string.Format("{0:N}", Convert.ToDecimal(Globals.oRec.Fields.Item("DocTotal").Value.ToString()));
                    if (Globals.oRec.Fields.Item("Nota").Value.ToString().Length > 0)
                    {
                        oitem.Nota = Globals.oRec.Fields.Item("Nota").Value.ToString();
                    }
                    oitem.ReferenciaSerie = Globals.oRec.Fields.Item("ReferenciaSerie").Value.ToString();

                    oitem.Sucursal = Globals.oRec.Fields.Item("Sucursal").Value.ToString();

                    oLISTA.Add(oitem);
                    Globals.oRec.MoveNext();
                }
                return oLISTA;
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        public static void ObtenerDatosCRE(int DocEntry, ref Retencion oDocumentoFE)
        {
            string nombre_Store = "EXX_RETENCION_FE";
            string nombre_Store1 = "EXX_RETENCION_DETALLE_FE";

            string Query = "";
            oDocumentoFE = new Retencion();

            try
            {

                #region SMC_Comprobante_21
                Query = "CALL \"" + nombre_Store + "\" (" + DocEntry + ")";
                Globals.oRec = Globals.RunQuery(Query);
                if (Globals.oRec != null)
                {
                    if (Globals.oRec.RecordCount > 0)
                    {
                        Globals.oRec.MoveFirst();
                        while (!Globals.oRec.EoF)
                        {
                            oDocumentoFE.correoEmisor = Globals.oRec.Fields.Item("correoEmisor").Value.ToString();
                            oDocumentoFE.correoAdquiriente = Globals.oRec.Fields.Item("correoAdquiriente").Value.ToString();
                            oDocumentoFE.serieNumeroRetencion = Globals.oRec.Fields.Item("serieNumeroRetencion").Value.ToString();
                            oDocumentoFE.fechaEmision = Globals.oRec.Fields.Item("fechaEmision").Value.ToString();
                            oDocumentoFE.tipoDocumento = Globals.oRec.Fields.Item("tipoDocumento").Value.ToString();
                            oDocumentoFE.numeroDocumentoEmisor = Globals.oRec.Fields.Item("numeroDocumentoEmisor").Value.ToString();
                            oDocumentoFE.tipoDocumentoEmisor = Globals.oRec.Fields.Item("tipoDocumentoEmisor").Value.ToString();
                            oDocumentoFE.razonSocialEmisor = Globals.oRec.Fields.Item("razonSocialEmisor").Value.ToString();
                            oDocumentoFE.numeroDocumentoEmisor = Globals.oRec.Fields.Item("numeroDocumentoEmisor").Value.ToString();
                            oDocumentoFE.nombreComercialEmisor = (Globals.oRec.Fields.Item("nombreComercialEmisor").Value.ToString());
                            oDocumentoFE.ubigeoEmisor = Globals.oRec.Fields.Item("ubigeoEmisor").Value.ToString();
                            oDocumentoFE.direccionEmisor = Globals.oRec.Fields.Item("direccionEmisor").Value.ToString();
                            oDocumentoFE.urbanizacionEmisor = Globals.oRec.Fields.Item("urbanizacionEmisor").Value.ToString();
                            oDocumentoFE.departamentoEmisor = Globals.oRec.Fields.Item("departamentoEmisor").Value.ToString();
                            oDocumentoFE.provinciaEmisor = Globals.oRec.Fields.Item("provinciaEmisor").Value.ToString();
                            oDocumentoFE.distritoEmisor = Globals.oRec.Fields.Item("distritoEmisor").Value.ToString();
                            oDocumentoFE.codigoPaisEmisor = Globals.oRec.Fields.Item("codigoPaisEmisor").Value.ToString();
                            oDocumentoFE.razonSocialProveedor = Globals.oRec.Fields.Item("razonSocialProveedor").Value.ToString();
                            oDocumentoFE.numeroDocumentoProveedor = Globals.oRec.Fields.Item("numeroDocumentoProveedor").Value.ToString();
                            oDocumentoFE.tipoDocumentoProveedor = Globals.oRec.Fields.Item("tipoDocumentoProveedor").Value.ToString();
                            oDocumentoFE.nombreComercialProveedor = (Globals.oRec.Fields.Item("nombreComercialProveedor").Value.ToString());
                            oDocumentoFE.direccionProveedor = Globals.oRec.Fields.Item("direccionProveedor").Value.ToString();
                            oDocumentoFE.urbanizacionProveedor = Globals.oRec.Fields.Item("urbanizacionProveedor").Value.ToString();
                            oDocumentoFE.provinciaProveedor = Globals.oRec.Fields.Item("provinciaProveedor").Value.ToString();
                            oDocumentoFE.departamentoProveedor = Globals.oRec.Fields.Item("departamentoProveedor").Value.ToString();
                            oDocumentoFE.distritoProveedor = Globals.oRec.Fields.Item("distritoProveedor").Value.ToString();
                            oDocumentoFE.codigoPaisProveedor = Globals.oRec.Fields.Item("codigoPaisProveedor").Value.ToString();
                            oDocumentoFE.ubigeoProveedor = Globals.oRec.Fields.Item("ubigeoProveedor").Value.ToString();
                            oDocumentoFE.regimenRetencion = Globals.oRec.Fields.Item("regimenRetencion").Value.ToString();

                            oDocumentoFE.tasaRetencion = (Globals.oRec.Fields.Item("tasaRetencion").Value.ToString());
                            oDocumentoFE.observaciones = (Globals.oRec.Fields.Item("observaciones").Value.ToString());
                            oDocumentoFE.importeTotalRetenido = (Globals.oRec.Fields.Item("importeTotalRetenido").Value.ToString());
                            oDocumentoFE.tipoMonedaTotalRetenido = (Globals.oRec.Fields.Item("tipoMonedaTotalRetenido").Value.ToString());

                            oDocumentoFE.importeTotalPagado = Globals.oRec.Fields.Item("importeTotalPagado").Value.ToString();
                            oDocumentoFE.tipoMonedaTotalPagado = Globals.oRec.Fields.Item("tipoMonedaTotalPagado").Value.ToString();



                            Globals.oRec.MoveNext();
                        }
                        Globals.Release(Globals.oRec);
                    }
                }
                #endregion

                #region SMC_COMPROBANTE_DETALLE_FE
                oDocumentoFE.RetencionItems = new List<RetencionItem>();
                RetencionItem oItem = new RetencionItem();
                Query = "CALL \"" + nombre_Store1 + "\" (" + DocEntry + ")";
                Globals.oRec = Globals.RunQuery(Query);
                if (Globals.oRec != null)
                {
                    if (Globals.oRec.RecordCount > 0)
                    {
                        int contadorItem = 1;
                        Globals.oRec.MoveFirst();
                        while (!Globals.oRec.EoF)
                        {
                            oItem = new RetencionItem();
                            oItem.numeroOrdenItem = contadorItem;
                            oItem.numeroDocumentoRelacionado = Globals.oRec.Fields.Item("numeroDocumentoRelacionado").Value.ToString();
                            oItem.fechaEmisionDocumentoRelacionado = (Globals.oRec.Fields.Item("fechaEmisionDocumentoRelacionado").Value.ToString());
                            oItem.tipoDocumentoRelacionado = Globals.oRec.Fields.Item("tipoDocumentoRelacionado").Value.ToString();
                            oItem.importeTotalDocumentoRelacionado = Globals.oRec.Fields.Item("importeTotalDocumentoRelacionado").Value.ToString();
                            oItem.tipoMonedaDocumentoRelacionado = Globals.oRec.Fields.Item("tipoMonedaDocumentoRelacionado").Value.ToString();
                            oItem.fechaPago = (Globals.oRec.Fields.Item("fechaPago").Value.ToString());
                            oItem.numeroPago = (Globals.oRec.Fields.Item("numeroPago").Value.ToString());
                            oItem.importePagoSinRetencion = Globals.oRec.Fields.Item("importePagoSinRetencion").Value.ToString();
                            oItem.importeRetenido = (Globals.oRec.Fields.Item("importeRetenido").Value.ToString());
                            oItem.monedaImporteRetenido = (Globals.oRec.Fields.Item("monedaImporteRetenido").Value.ToString());
                            oItem.monedaPago = (Globals.oRec.Fields.Item("monedaPago").Value.ToString());

                            oItem.fechaRetencion = (Globals.oRec.Fields.Item("fechaRetencion").Value.ToString());
                            oItem.importeTotalPagarNeto = (Globals.oRec.Fields.Item("importeTotalPagarNeto").Value.ToString());
                            oItem.monedaMontoNetoPagado = (Globals.oRec.Fields.Item("monedaMontoNetoPagado").Value.ToString());


                            if (Globals.oRec.Fields.Item("monedaPago").Value.ToString() != "PEN")
                            {
                                oItem.monedaReferenciaTipoCambio = (Globals.oRec.Fields.Item("monedaReferenciaTipoCambio").Value.ToString());
                                oItem.monedaObjetivoTasaCambio = (Globals.oRec.Fields.Item("monedaObjetivoTasaCambio").Value.ToString());
                                oItem.factorTipoCambioMoneda = (Globals.oRec.Fields.Item("factorTipoCambioMoneda").Value.ToString());
                                oItem.fechaCambio = (Globals.oRec.Fields.Item("fechaCambio").Value.ToString());
                            }


                            contadorItem++;
                            oDocumentoFE.RetencionItems.Add(oItem);
                            Globals.oRec.MoveNext();
                        }
                    }
                }
                #endregion

                #region Calculando Importes Totales
                decimal SumimporteTotalPagado = 0;
                decimal SumimporteTotalRetenido = 0;
                foreach (var item in oDocumentoFE.RetencionItems)
                {
                    SumimporteTotalPagado += Convert.ToDecimal(item.importeTotalPagarNeto);
                    SumimporteTotalRetenido += Convert.ToDecimal(item.importeRetenido);
                    //oDocumentoFE.importeTotalPagado += Convert.ToDecimal(item.importeTotalPagarNeto);
                    //oDocumentoFE.importeTotalRetenido += Convert.ToDecimal(item.importeRetenido);
                }
                oDocumentoFE.importeTotalPagado = SumimporteTotalPagado.ToString();
                oDocumentoFE.importeTotalRetenido = SumimporteTotalRetenido.ToString();
                #endregion
            }
            catch (Exception ex)
            {
                ex.Message.ToString();
            }
            finally
            {
                Globals.Release(Globals.oRec);
                GC.Collect();
            }
        }

        public static void ObtenerDatosCPEPaperless(ref Entities.Paperless.Documento oBeFEDOC, string Table)
        {
            try
            {
                Globals.Query = ServicioConectorFE.Properties.Resources.ObtenerDatosCPE;
                Globals.Query = string.Format(Globals.Query, Table, oBeFEDOC.DocEntry.ToString(), Table.Substring(1, 3));
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();

                if (Globals.oRec.RecordCount > 0)
                {
                    oBeFEDOC.documentoExportacion = Globals.oRec.Fields.Item("documentoExportacion").Value.ToString();
                    oBeFEDOC.TransGratuita = Globals.oRec.Fields.Item("TransGratuita").Value.ToString() == "1" ? false : true;
                    oBeFEDOC.CardCode = Globals.oRec.Fields.Item("CardCode").Value.ToString();
                    oBeFEDOC.PayToCode = Globals.oRec.Fields.Item("PayToCode").Value.ToString();
                    oBeFEDOC.ShipToCode = Globals.oRec.Fields.Item("ShipToCode").Value.ToString();
                    oBeFEDOC.DocType = Globals.oRec.Fields.Item("DocType").Value.ToString();
                    oBeFEDOC.tipoComprobante = Globals.oRec.Fields.Item("tipoComprobante").Value.ToString();
                    oBeFEDOC.serie = Globals.oRec.Fields.Item("serie").Value.ToString();
                    oBeFEDOC.correlativo = Globals.oRec.Fields.Item("correlativo").Value.ToString().PadLeft(8, '0');
                    oBeFEDOC.fechaEmision = Globals.oRec.Fields.Item("fechaEmision").Value.ToString();
                    oBeFEDOC.tipoMoneda = Globals.oRec.Fields.Item("tipoMoneda").Value.ToString();
                    oBeFEDOC.TipoCambio = Globals.oRec.Fields.Item("TipoCambio").Value.ToString();
                    oBeFEDOC.tipoOperacion = Globals.oRec.Fields.Item("tipoOperacion").Value.ToString();
                    oBeFEDOC.ordenCompra = Globals.oRec.Fields.Item("ordenCompra").Value.ToString();
                    oBeFEDOC.redondeo = Globals.RetirarMonedayDecimal(Globals.oRec.Fields.Item("redondeo").Value.ToString());
                    oBeFEDOC.fechaVencimiento = Globals.oRec.Fields.Item("fechaVencimiento").Value.ToString();
                    oBeFEDOC.horaEmision = Globals.oRec.Fields.Item("horaEmision").Value.ToString();
                    oBeFEDOC.codigoEstablecimiento = Globals.oRec.Fields.Item("codigoEstablecimiento").Value.ToString();
                    oBeFEDOC.pedido = ""; // Globals.oRec.Fields.Item("pedido").Value.ToString();
                    oBeFEDOC.observaciones = Globals.oRec.Fields.Item("observaciones").Value.ToString();
                    oBeFEDOC.vendedor = Globals.oRec.Fields.Item("vendedor").Value.ToString();
                    oBeFEDOC.condicionPago = Globals.oRec.Fields.Item("condicionPago").Value.ToString();
                    oBeFEDOC.condicionPagoDesc = Globals.oRec.Fields.Item("condicionPagoDesc").Value.ToString();
                    oBeFEDOC.importeTotal = Globals.RetirarMonedayDecimal(Globals.oRec.Fields.Item("importeTotal").Value.ToString());
                    oBeFEDOC.valorTotalAnticipos = Globals.RetirarMonedayDecimal(Globals.oRec.Fields.Item("valorTotalAnticipos").Value.ToString());
                    oBeFEDOC.valorTotalIGVAnticipos = Globals.RetirarMonedayDecimal(Globals.oRec.Fields.Item("valorTotalIGVAnticipos").Value.ToString());
                    oBeFEDOC.valorDescuentoGlobal = Globals.RetirarMonedayDecimal(Globals.oRec.Fields.Item("valorDescuentoGlobal").Value.ToString());
                    oBeFEDOC.Comments = Globals.oRec.Fields.Item("Comments").Value.ToString();
                    oBeFEDOC.credito = Globals.oRec.Fields.Item("Credito").Value.ToString();
                    oBeFEDOC.TieneDetraccion = Globals.oRec.Fields.Item("TieneDetraccion").Value.ToString();
                    oBeFEDOC.codigoDetraccion = Globals.oRec.Fields.Item("codigoDetraccion").Value.ToString();
                    oBeFEDOC.totalDetraccion = Globals.oRec.Fields.Item("totalDetraccion").Value.ToString();
                    oBeFEDOC.totalDetraccionMD = Globals.oRec.Fields.Item("totalDetraccionMD").Value.ToString();
                    oBeFEDOC.porcentajeDetraccion = Globals.oRec.Fields.Item("porcentajeDetraccion").Value.ToString();

                    oBeFEDOC.notaTipo = "";
                    oBeFEDOC.notaSustento = "";
                    oBeFEDOC.ListReferencia = new List<Entities.Paperless.Referencia>();

                    if (oBeFEDOC.tipoComprobante == "07" || oBeFEDOC.tipoComprobante == "08")
                    {
                        oBeFEDOC.notaTipo = Globals.oRec.Fields.Item("notaTipo").Value.ToString();
                        oBeFEDOC.notaSustento = Globals.oRec.Fields.Item("notaSustento").Value.ToString();

                        Entities.Paperless.Referencia oReferencia = new Entities.Paperless.Referencia();
                        oReferencia.tipoDocumento = Globals.oRec.Fields.Item("tipoDocumentoRef").Value.ToString();
                        oReferencia.serie = Globals.oRec.Fields.Item("serieRef").Value.ToString();
                        oReferencia.correlativo = Globals.oRec.Fields.Item("correlativoRef").Value.ToString();
                        oReferencia.fechaEmision = Globals.oRec.Fields.Item("fechaEmisionRef").Value.ToString();
                        oBeFEDOC.ListReferencia.Add(oReferencia);
                    }

                    Globals.Release(Globals.oRec);

                    //Datos de Emisor
                    oBeFEDOC.emisor = new Entities.Paperless.Emisor();
                    oBeFEDOC.emisor.ruc = Globals.oOADM.TaxIdNum;
                    oBeFEDOC.emisor.tipoIdentificacion = "6";
                    oBeFEDOC.emisor.razonSocial = Globals.oOADM.CompnyName;
                    oBeFEDOC.emisor.nombreComercial = Globals.oOADM.CompnyName;
                    oBeFEDOC.emisor.ubigeo = Globals.oADM1.ZipCode;
                    oBeFEDOC.emisor.direccion = Globals.oADM1.Street;
                    oBeFEDOC.emisor.departamento = Globals.oADM1.State;
                    oBeFEDOC.emisor.provincia = Globals.oADM1.County;
                    oBeFEDOC.emisor.distrito = Globals.oADM1.City;

                    //Datos de Adquiriente
                    Globals.Query = ServicioConectorFE.Properties.Resources.ObtenerDatosAdquiriente;
                    Globals.Query = string.Format(Globals.Query, oBeFEDOC.CardCode, oBeFEDOC.PayToCode, oBeFEDOC.ShipToCode);
                    Globals.RunQuery(Globals.Query);
                    Globals.oRec.MoveFirst();

                    if (Globals.oRec.RecordCount == 0)
                    {
                        throw new Exception("No se encuentra datos de Adquiriente.");
                    }
                    else
                    {
                        oBeFEDOC.adquiriente = new Entities.Paperless.Adquiriente();
                        oBeFEDOC.adquiriente.tipoIdentificacion = Globals.oRec.Fields.Item("tipoIdentificacion").Value.ToString().Trim();
                        oBeFEDOC.adquiriente.numeroIdentificacion = Globals.oRec.Fields.Item("numeroIdentificacion").Value.ToString().Trim();
                        oBeFEDOC.adquiriente.nombre = Globals.oRec.Fields.Item("nombre").Value.ToString().Trim().ToUpper();
                        oBeFEDOC.adquiriente.direccionFiscal = Globals.oRec.Fields.Item("direccionFiscal").Value.ToString().Trim().ToUpper();
                        oBeFEDOC.adquiriente.email = Globals.oRec.Fields.Item("email").Value.ToString().Trim().ToUpper().Replace("\"", "");
                        //oBEiFEDoc.adquiriente.email = "ivan.marco@mixercon.com";
                        oBeFEDOC.adquiriente.codigoPais = Globals.oRec.Fields.Item("codigoPais").Value.ToString().Trim().ToUpper();
                        oBeFEDOC.adquiriente.telefono = Globals.oRec.Fields.Item("telefono").Value.ToString().Trim().ToUpper();
                        oBeFEDOC.adquiriente.fax = Globals.oRec.Fields.Item("fax").Value.ToString().Trim().ToUpper();
                        oBeFEDOC.adquiriente.obra = Globals.oRec.Fields.Item("obra").Value.ToString().Trim().ToUpper();
                        Globals.Release(Globals.oRec);
                    }


                    //Datos Adicionales
                    Globals.Query = ServicioConectorFE.Properties.Resources.ObtenerDatosAdicionalesPaperless;
                    Globals.Query = string.Format(Globals.Query, Table, oBeFEDOC.DocEntry.ToString(), Table.Substring(1, 3));
                    Globals.RunQuery(Globals.Query);
                    Globals.oRec.MoveFirst();

                    if (Globals.oRec.RecordCount > 0)
                    {
                        oBeFEDOC.U_EXT_NAVE = Globals.oRec.Fields.Item("U_EXT_NAVE").Value.ToString();
                        oBeFEDOC.U_EXT_CALLNUMBER = Globals.oRec.Fields.Item("U_EXT_CALLNUMBER").Value.ToString();
                        oBeFEDOC.U_EXT_DESRESOLIDA = Globals.oRec.Fields.Item("U_EXT_DESRESOLIDA").Value.ToString();
                        oBeFEDOC.U_EXT_OSNRO = Globals.oRec.Fields.Item("U_EXT_OSNRO").Value.ToString();
                        oBeFEDOC.U_EXT_COMENTARIO = Globals.oRec.Fields.Item("U_EXT_COMENTARIO").Value.ToString();
                        oBeFEDOC.U_EXT_MUELLE = Globals.oRec.Fields.Item("U_EXT_MUELLE").Value.ToString();
                        oBeFEDOC.U_EXT_VIAJE = Globals.oRec.Fields.Item("U_EXT_VIAJE").Value.ToString();
                        oBeFEDOC.U_EXT_FECARRIBO = Globals.oRec.Fields.Item("U_EXT_FECARRIBO").Value.ToString();
                        oBeFEDOC.U_EXT_LINEANAV = Globals.oRec.Fields.Item("U_EXT_LINEANAV").Value.ToString();
                        oBeFEDOC.U_EXT_RESPSOLIDA = Globals.oRec.Fields.Item("U_EXT_RESPSOLIDA").Value.ToString();
                        oBeFEDOC.U_EXT_dc_fact_VBWONbr = Globals.oRec.Fields.Item("U_EXT_dc_fact_VBWONbr").Value.ToString();
                        oBeFEDOC.U_EXT_FECARRIBO = string.IsNullOrEmpty(oBeFEDOC.U_EXT_FECARRIBO) ? "" : Convert.ToDateTime(oBeFEDOC.U_EXT_FECARRIBO).ToString("dd/MM/yyyy");

                        Globals.Release(Globals.oRec);
                    }
                }
                else
                {
                    throw new Exception("No se logró obtener datos.");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }


        public static void ObtenerDatosCPEEstela(ref Entities.Estela.Documento oBeFEDOC, string Table)
        {
            try
            {
                Globals.Query = ServicioConectorFE.Properties.Resources.ObtenerDatosCPE;
                Globals.Query = string.Format(Globals.Query, Table, oBeFEDOC.DocEntry.ToString(), Table.Substring(1, 3));
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();

                if (Globals.oRec.RecordCount > 0)
                {
                    oBeFEDOC.documentoExportacion = Globals.oRec.Fields.Item("documentoExportacion").Value.ToString();
                    oBeFEDOC.TransGratuita = Globals.oRec.Fields.Item("TransGratuita").Value.ToString() == "1" ? false : true;
                    oBeFEDOC.CardCode = Globals.oRec.Fields.Item("CardCode").Value.ToString();
                    oBeFEDOC.PayToCode = Globals.oRec.Fields.Item("PayToCode").Value.ToString();
                    oBeFEDOC.ShipToCode = Globals.oRec.Fields.Item("ShipToCode").Value.ToString();
                    oBeFEDOC.DocType = Globals.oRec.Fields.Item("DocType").Value.ToString();
                    oBeFEDOC.tipoComprobante = Globals.oRec.Fields.Item("tipoComprobante").Value.ToString();
                    oBeFEDOC.serie = Globals.oRec.Fields.Item("serie").Value.ToString();
                    oBeFEDOC.correlativo = Globals.oRec.Fields.Item("correlativo").Value.ToString().PadLeft(8, '0');
                    oBeFEDOC.fechaEmision = Globals.oRec.Fields.Item("fechaEmision").Value.ToString();
                    oBeFEDOC.tipoMoneda = Globals.oRec.Fields.Item("tipoMoneda").Value.ToString();
                    oBeFEDOC.TipoCambio = Globals.oRec.Fields.Item("TipoCambio").Value.ToString();
                    oBeFEDOC.tipoOperacion = Globals.oRec.Fields.Item("tipoOperacion").Value.ToString();
                    oBeFEDOC.ordenCompra = Globals.oRec.Fields.Item("ordenCompra").Value.ToString();
                    oBeFEDOC.redondeo = Globals.RetirarMonedayDecimal(Globals.oRec.Fields.Item("redondeo").Value.ToString());
                    oBeFEDOC.fechaVencimiento = Globals.oRec.Fields.Item("fechaVencimiento").Value.ToString();
                    oBeFEDOC.horaEmision = Globals.oRec.Fields.Item("horaEmision").Value.ToString();
                    oBeFEDOC.codigoEstablecimiento = Globals.oRec.Fields.Item("codigoEstablecimiento").Value.ToString();
                    oBeFEDOC.pedido = ""; // Globals.oRec.Fields.Item("pedido").Value.ToString();
                    oBeFEDOC.observaciones = Globals.oRec.Fields.Item("observaciones").Value.ToString();
                    oBeFEDOC.vendedor = Globals.oRec.Fields.Item("vendedor").Value.ToString();
                    oBeFEDOC.condicionPago = Globals.oRec.Fields.Item("condicionPago").Value.ToString();
                    oBeFEDOC.condicionPagoDesc = Globals.oRec.Fields.Item("condicionPagoDesc").Value.ToString();
                    oBeFEDOC.importeTotal = Globals.RetirarMonedayDecimal(Globals.oRec.Fields.Item("importeTotal").Value.ToString());
                    oBeFEDOC.valorTotalAnticipos = Globals.RetirarMonedayDecimal(Globals.oRec.Fields.Item("valorTotalAnticipos").Value.ToString());
                    oBeFEDOC.valorTotalIGVAnticipos = Globals.RetirarMonedayDecimal(Globals.oRec.Fields.Item("valorTotalIGVAnticipos").Value.ToString());
                    oBeFEDOC.valorDescuentoGlobal = Globals.RetirarMonedayDecimal(Globals.oRec.Fields.Item("valorDescuentoGlobal").Value.ToString());
                    oBeFEDOC.Comments = Globals.oRec.Fields.Item("Comments").Value.ToString();
                    oBeFEDOC.credito = Globals.oRec.Fields.Item("Credito").Value.ToString();
                    oBeFEDOC.TieneDetraccion = Globals.oRec.Fields.Item("TieneDetraccion").Value.ToString();
                    oBeFEDOC.codigoDetraccion = Globals.oRec.Fields.Item("codigoDetraccion").Value.ToString();
                    oBeFEDOC.totalDetraccion = Globals.oRec.Fields.Item("totalDetraccion").Value.ToString();
                    oBeFEDOC.totalDetraccion = Math.Round(Convert.ToDouble(oBeFEDOC.totalDetraccion), 0).ToString();
                    oBeFEDOC.totalDetraccionMD = Globals.oRec.Fields.Item("totalDetraccionMD").Value.ToString();
                    oBeFEDOC.porcentajeDetraccion = Globals.oRec.Fields.Item("porcentajeDetraccion").Value.ToString();
                    oBeFEDOC.porcentajeDetraccion = Math.Round((Convert.ToDouble(oBeFEDOC.porcentajeDetraccion) * 100), 0).ToString();
                    oBeFEDOC.Responsable = Globals.oRec.Fields.Item("Responsable").Value.ToString();

                    oBeFEDOC.notaTipo = "";
                    oBeFEDOC.notaSustento = "";
                    oBeFEDOC.ListReferencia = new List<Entities.Estela.Referencia>();

                    if (oBeFEDOC.tipoComprobante == "07" || oBeFEDOC.tipoComprobante == "08")
                    {
                        oBeFEDOC.notaTipo = Globals.oRec.Fields.Item("notaTipo").Value.ToString();
                        oBeFEDOC.notaSustento = Globals.oRec.Fields.Item("notaSustento").Value.ToString();

                        Entities.Estela.Referencia oReferencia = new Entities.Estela.Referencia();
                        oReferencia.tipoDocumento = Globals.oRec.Fields.Item("tipoDocumentoRef").Value.ToString();
                        oReferencia.serie = Globals.oRec.Fields.Item("serieRef").Value.ToString();
                        oReferencia.correlativo = Globals.oRec.Fields.Item("correlativoRef").Value.ToString();
                        oReferencia.fechaEmision = Globals.oRec.Fields.Item("fechaEmisionRef").Value.ToString();
                        oBeFEDOC.ListReferencia.Add(oReferencia);

                        //Datos Adicionales
                        Globals.Query = ServicioConectorFE.Properties.Resources.ObtenerDatosAdicionalesEstela;
                        Globals.Query = string.Format(Globals.Query, Table, oBeFEDOC.DocEntry.ToString(), Table.Substring(1, 3));
                        Globals.RunQuery(Globals.Query);
                        Globals.oRec.MoveFirst();

                        if (Globals.oRec.RecordCount > 0)
                        {
                            oBeFEDOC.notaTipo = Globals.oRec.Fields.Item("U_EXA_TIPNCND").Value.ToString();
                            oBeFEDOC.notaSustento = Globals.oRec.Fields.Item("U_EXA_DESMOT").Value.ToString();

                            Globals.Release(Globals.oRec);
                        }
                    }

                    Globals.Release(Globals.oRec);


                    //Datos de Emisor
                    oBeFEDOC.emisor = new Entities.Estela.Emisor();
                    oBeFEDOC.emisor.ruc = Globals.oOADM.TaxIdNum;
                    oBeFEDOC.emisor.tipoIdentificacion = "6";
                    oBeFEDOC.emisor.razonSocial = Globals.oOADM.PrintHeader;
                    oBeFEDOC.emisor.nombreComercial = Globals.oOADM.PrintHeader;
                    oBeFEDOC.emisor.ubigeo = Globals.oADM1.ZipCode;
                    oBeFEDOC.emisor.direccion = Globals.oADM1.Street;
                    oBeFEDOC.emisor.departamento = Globals.oADM1.State;
                    oBeFEDOC.emisor.provincia = Globals.oADM1.County;
                    oBeFEDOC.emisor.distrito = Globals.oADM1.City;

                    //Datos de Adquiriente
                    Globals.Query = ServicioConectorFE.Properties.Resources.ObtenerDatosAdquiriente;
                    Globals.Query = string.Format(Globals.Query, oBeFEDOC.CardCode, oBeFEDOC.PayToCode, oBeFEDOC.ShipToCode);
                    Globals.RunQuery(Globals.Query);
                    Globals.oRec.MoveFirst();

                    if (Globals.oRec.RecordCount == 0)
                    {
                        throw new Exception("No se encuentra datos de Adquiriente.");
                    }
                    else
                    {
                        oBeFEDOC.adquiriente = new Adquiriente();
                        oBeFEDOC.adquiriente.tipoIdentificacion = Globals.oRec.Fields.Item("tipoIdentificacion").Value.ToString().Trim();
                        oBeFEDOC.adquiriente.numeroIdentificacion = Globals.oRec.Fields.Item("numeroIdentificacion").Value.ToString().Trim();
                        oBeFEDOC.adquiriente.nombre = Globals.oRec.Fields.Item("nombre").Value.ToString().Trim().ToUpper();
                        oBeFEDOC.adquiriente.ubigeo = Globals.oRec.Fields.Item("GlblLocNum").Value.ToString().Trim().ToUpper();
                        oBeFEDOC.adquiriente.direccionFiscal = Globals.oRec.Fields.Item("direccionFiscal").Value.ToString().Trim().ToUpper();
                        oBeFEDOC.adquiriente.email = Globals.oRec.Fields.Item("email").Value.ToString().Trim().ToUpper().Replace("\"", "");
                        //oBEiFEDoc.adquiriente.email = "ivan.marco@mixercon.com";
                        oBeFEDOC.adquiriente.codigoPais = Globals.oRec.Fields.Item("codigoPais").Value.ToString().Trim().ToUpper();
                        oBeFEDOC.adquiriente.telefono = Globals.oRec.Fields.Item("telefono").Value.ToString().Trim().ToUpper();
                        oBeFEDOC.adquiriente.fax = Globals.oRec.Fields.Item("fax").Value.ToString().Trim().ToUpper();
                        oBeFEDOC.adquiriente.obra = Globals.oRec.Fields.Item("obra").Value.ToString().Trim().ToUpper();
                        Globals.Release(Globals.oRec);
                    }
                }
                else
                {
                    throw new Exception("No se logró obtener datos.");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static void ObtenerDetalleDocumento(ref DataTable dt, string SAPTable, int docEntry)
        {
            try
            {
                if (Globals.ProveedorOSE == "3")
                    Globals.Query = ServicioConectorFE.Properties.Resources.ObtenerDetalleEstela;
                else
                    Globals.Query = ServicioConectorFE.Properties.Resources.ObtenerDetalleDocumento;

                Globals.Query = String.Format(Globals.Query, SAPTable, SAPTable.Substring(1, 3), docEntry);
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();

                if (Globals.oRec.RecordCount == 0)
                {
                    throw new Exception("No se encuentra datos de documentos referenciados del documento: " + SAPTable + "-" + docEntry);
                }
                else
                {
                    dt = Globals.RecorsetToDataTable(Globals.oRec);
                    dt.TableName = "dtDetalle";
                    Globals.Release(Globals.oRec);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                Globals.Release(Globals.oRec);
            }
        }

        public static bool AnexarDocumentosCpe(int DocEntry, string nombreArchivo, string pdf_url, string xml_url, string cdr_url, int ObjType, SAPbobsCOM.Company oCompany, ref string mensaje)
        {
            Funciones oFuncionesRequeridas = new Funciones();

            #region PDF
            string mensajeerrorpdf = "";


            if (!ValidarPesoArchivo(nombreArchivo + ".pdf"))
            {
                string rutaArchivoPdf = Globals.PathBackup + nombreArchivo + ".pdf";
                if (!string.IsNullOrEmpty(pdf_url))
                {
                    if (System.IO.File.Exists(rutaArchivoPdf))
                        System.IO.File.Delete(rutaArchivoPdf);

                    if (Globals.ProveedorOSE == "3")
                    {
                        byte[] pdfBytes = Convert.FromBase64String(pdf_url);
                        System.IO.File.WriteAllBytes(rutaArchivoPdf, pdfBytes);
                    }
                    else
                        oFuncionesRequeridas.downloadFileToSpecificPath(pdf_url, rutaArchivoPdf);
                }
                SetAnexoSAPCPE(nombreArchivo, "pdf", Globals.PathBackup, DocEntry, oCompany, ObjType, ref mensajeerrorpdf);
            }

            #endregion

            #region XML
            if (!ValidarPesoArchivo(nombreArchivo + ".xml"))
            {
                string rutaArchivoXml = Globals.PathBackup + nombreArchivo + ".xml";
                if (!string.IsNullOrEmpty(xml_url))
                {
                    if (System.IO.File.Exists(rutaArchivoXml))
                        System.IO.File.Delete(rutaArchivoXml);

                    if (Globals.ProveedorOSE == "3")
                        System.IO.File.WriteAllText(rutaArchivoXml, xml_url);
                    else

                        oFuncionesRequeridas.downloadFileToSpecificPath(xml_url, rutaArchivoXml);
                }
                SetAnexoSAPCPE(nombreArchivo, "xml", Globals.PathBackup, DocEntry, oCompany, ObjType, ref mensajeerrorpdf);

            }
            #endregion

            #region CDR
            if (!ValidarPesoArchivo("R-" + nombreArchivo + ".xml"))
            {
                string rutaArchivoCdr = Globals.PathBackup + "R-" + nombreArchivo + ".xml";
                if (!string.IsNullOrEmpty(cdr_url))
                {
                    if (System.IO.File.Exists(rutaArchivoCdr))
                        System.IO.File.Delete(rutaArchivoCdr);

                    if (Globals.ProveedorOSE == "3")
                        System.IO.File.WriteAllText(rutaArchivoCdr, cdr_url);
                    if (Globals.ProveedorOSE == "2")
                    {
                        byte[] pdfBytes = Convert.FromBase64String(cdr_url);
                        System.IO.File.WriteAllBytes(rutaArchivoCdr, pdfBytes);
                    }
                    else
                        oFuncionesRequeridas.downloadFileToSpecificPath(cdr_url, rutaArchivoCdr);
                }
                SetAnexoSAPCPE("R-" + nombreArchivo, "xml", Globals.PathBackup, DocEntry, oCompany, ObjType, ref mensajeerrorpdf);
            }
            #endregion

            return true;
        }

        public static bool ValidarPesoArchivo(string NombreArchivo)
        {
            bool respuesta = false;
            SAPbobsCOM.CompanyService comServ = Globals.oCompany.GetCompanyService();
            SAPbobsCOM.PathAdmin pthAdm = comServ.GetPathAdmin();
            string path = pthAdm.AttachmentsFolderPath;
            if (path.Substring(path.Length - 1, 1).Equals(@"\"))
                path = path.Substring(0, path.Length - 1);

            string rutaArchivo = path + "\\" + NombreArchivo;
            // Verificar si el archivo existe
            if (System.IO.File.Exists(rutaArchivo))
            {
                // Obtener información del archivo
                FileInfo fileInfo = new FileInfo(rutaArchivo);

                // Obtener el tamaño del archivo en bytes
                long pesoEnBytes = fileInfo.Length;

                // Convertir el tamaño a kilobytes (1 KB = 1024 bytes)
                double pesoEnKilobytes = (double)pesoEnBytes / 1024;
                if (pesoEnKilobytes > 0)
                {
                    respuesta = true;
                }

            }
            return respuesta;
        }

        public static bool SetAnexoSAPCPE(string nombrearchivo, string extension, string rutaarchivo, int docentry, SAPbobsCOM.Company oCompany, int ObjectType, ref string mensajeErrorInterno)
        {
            int existearchivoanexado = 0;
            int numeroexistencia = 0;
            try
            {
                string exportPathForSAP = rutaarchivo;
                exportPathForSAP = exportPathForSAP.Replace("\\\\\\\\", "\\\\");
                exportPathForSAP = exportPathForSAP.Replace("XML\\", "XML");

                SAPbobsCOM.Payments opay = null;
                SAPbobsCOM.Documents oinvoice = null;
                if (ObjectType == 46)
                {
                    opay = (SAPbobsCOM.Payments)oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Convert.ToInt32(ObjectType));
                    opay.GetByKey(Convert.ToInt32(docentry));
                }
                else
                {
                    oinvoice = (SAPbobsCOM.Documents)oCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Convert.ToInt32(ObjectType));
                    oinvoice.GetByKey(Convert.ToInt32(docentry));
                }

                SAPbobsCOM.Attachments2 oAttachment = (SAPbobsCOM.Attachments2)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2);
                SAPbobsCOM.CompanyService comServ = oCompany.GetCompanyService();
                SAPbobsCOM.PathAdmin pthAdm = comServ.GetPathAdmin();
                string path = pthAdm.AttachmentsFolderPath;
                if (path.Substring(path.Length - 1, 1).Equals(@"\"))
                    path = path.Substring(0, path.Length - 1);

                if (oAttachment.GetByKey(ObjectType == 46 ? opay.AttachmentEntry : oinvoice.AttachmentEntry))
                {
                    for (int i = 0; i < oAttachment.Lines.Count; i++)
                    {
                        oAttachment.Lines.SetCurrentLine(i);
                        if (oAttachment.Lines.FileName == nombrearchivo && oAttachment.Lines.FileExtension == extension)
                        {
                            existearchivoanexado = 1;
                            numeroexistencia = i;
                        }
                        oAttachment.Lines.SourcePath = path;
                    }
                    oAttachment.Lines.Add();
                    int sadasdasdasdas = oAttachment.Lines.Count;
                    oAttachment.Lines.SetCurrentLine(oAttachment.Lines.Count - 1);
                    if (existearchivoanexado == 1)
                    {
                        oAttachment.Lines.SetCurrentLine(numeroexistencia);
                        EliminiarArchivoSAPAnexo(exportPathForSAP + nombrearchivo, nombrearchivo, extension);
                    }

                    oAttachment.Lines.FileName = nombrearchivo;
                    oAttachment.Lines.FileExtension = extension;
                    oAttachment.Lines.SourcePath = exportPathForSAP;
                    oAttachment.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES;

                    if (oAttachment.Update() != 0)
                    {
                        mensajeErrorInterno = oCompany.GetLastErrorDescription();

                        if (mensajeErrorInterno.Length > 0)
                        {
                            oAttachment.Lines.SetCurrentLine(oAttachment.Lines.Count - 1);
                            oAttachment.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES;
                            oAttachment.Lines.FileName = oAttachment.Lines.FileName;
                            oAttachment.Lines.SourcePath = exportPathForSAP.Replace("WIN-SOPORTESAP", "192.168.1.9");

                            if (oAttachment.Update() != 0)
                                mensajeErrorInterno = oCompany.GetLastErrorDescription();
                        }
                    }
                }
                else
                {
                    oAttachment.Lines.FileName = nombrearchivo;
                    oAttachment.Lines.FileExtension = extension;
                    oAttachment.Lines.SourcePath = exportPathForSAP;
                    oAttachment.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES;
                    if (oAttachment.Add() != 0)
                        mensajeErrorInterno = oCompany.GetLastErrorDescription();

                    string objKey = oCompany.GetNewObjectKey();
                    oAttachment.GetByKey(Convert.ToInt32(objKey));
                    int absEntry = oAttachment.AbsoluteEntry;

                    if (ObjectType == 46)
                    {
                        opay.AttachmentEntry = oAttachment.AbsoluteEntry;

                        if (opay.Update() != 0)
                            mensajeErrorInterno = oCompany.GetLastErrorDescription();
                    }
                    else
                    {
                        oinvoice.AttachmentEntry = oAttachment.AbsoluteEntry;

                        if (oinvoice.Update() != 0)
                            mensajeErrorInterno = oCompany.GetLastErrorDescription();
                    }
                }
            }
            catch (System.Exception ex)
            {
                mensajeErrorInterno = ex.Message;
            }

            if (mensajeErrorInterno.Length > 1)
                return false;

            return true;
        }

        public static bool EliminiarArchivoSAPAnexo(string RutaArchivo, string NombreArchivo, string extension)
        {
            bool respuesta = false;

            #region Eliminar Archivo
            try
            {
                RutaArchivo = RutaArchivo + "." + extension;
                if (System.IO.File.Exists(RutaArchivo))
                {
                    // Borrar el archivo
                    System.IO.File.Delete(RutaArchivo);
                    respuesta = true;
                }
                else
                {
                    respuesta = false;
                }
            }
            catch (System.Exception ex)
            {
                // Capturar y manejar cualquier excepción que pueda ocurrir durante el proceso de borrado
                respuesta = false;
            }
            #endregion

            #region Copiar Archivo
            try
            {
                string rutaoriginal = Globals.PathBackup + NombreArchivo + "." + extension;
                // Verificar si el archivo existe antes de intentar borrarlo
                if (System.IO.File.Exists(rutaoriginal))
                {
                    // Borrar el archivo
                    System.IO.File.Copy(rutaoriginal, RutaArchivo);
                }
            }
            catch (System.Exception ex)
            {
            }
            #endregion

            return respuesta;

        }

        public static void ActualizarEstadoFE(int ObjectType, int DocEntry, string MensajeError, string Estado, SAPbobsCOM.Company oCompanyDAO)
        {
            string nombre_Store = "EXX_ActualizarEstadoFE";
            SAPbobsCOM.Recordset oRecFE = default(SAPbobsCOM.Recordset);
            string Query = "";
            try
            {
                string mensaje = MensajeError.Replace("'", "").Replace("\"", "").Replace("\r\n", "").Replace("\n", "").Replace("\r", "");
                mensaje = mensaje.Length > 148 ? mensaje.Substring(0, 148) : mensaje;
                Query = "CALL \"" + nombre_Store + "\" ('" + DocEntry + "','" + ObjectType + "','" + mensaje + "','" + Estado + "')";
                oRecFE = Globals.RunQuery(Query);

            }
            catch (Exception ex)
            {
                ex.Message.ToString();
            }
            finally
            {
                Globals.Release(oRecFE);
                oRecFE = null;
                GC.Collect();
            }
        }
    }
}
