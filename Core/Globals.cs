using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ServicioConectorFE.Entities.SAP;
using System.Data;
using ServicioConectorFE.Framework;
using SAPbobsCOM;
using ServicioConectorFE.Entities.Estela;
using ServicioConectorFE.OSE;
using System.Diagnostics;

namespace ServicioConectorFE.Core
{
    public class Globals
    {
        public static String AddOnName = "Servicio Conector FE";
        public static String AddOnVersion = "1.0.0";

        public static string configPath = AppDomain.CurrentDomain.BaseDirectory + "SetupFields\\";
        public static List<CONF> ListConfig = new List<CONF>();
        public static List<string> oDocuments = new List<string> { "01", "03", "07", "08", "20" };
        public static System.Data.DataTable dtMonedas;
        public static Entities.SAP.OADM oOADM = new Entities.SAP.OADM();
        public static Entities.SAP.ADM1 oADM1 = new Entities.SAP.ADM1();

        public static string URL_OSE;
        public static string SFTP_IP;
        public static int SFTP_PORT;
        public static string SFTP_USER;
        public static string SFTP_PASS;

        public static string ProveedorOSE;
        public static string PathArchivos;
        public static string PathBackup = @"C:\archivos txt\";
        public static Double TipoCambio = 0;

        public static string IdEmpresa;
        public static SAPbobsCOM.Company oCompany;
        public static SAPbobsCOM.Recordset oRec = default(SAPbobsCOM.Recordset);
        public static string Query;

        #region FuncionesGenerales
        public static List<T> ReadJson<T>(string FileName) where T : class, new()
        {
            try
            {
                if (System.IO.File.Exists(configPath + FileName))
                {
                    string jsonContent = System.IO.File.ReadAllText(configPath + FileName);
                    JArray contentArray = JArray.Parse(jsonContent);
                    return contentArray.ToObject<List<T>>();
                }
                else
                    throw new Exception($"Archivo {FileName} no encontrado.");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        public static void GetConfig()
        {
            try
            {
                ListConfig = Globals.ReadJson<CONF>("Config.json");

                ListConfig.ForEach(x =>
                {
                    x = new CONF
                    {
                        IdEmpresa = x.IdEmpresa,
                        DbServerType = x.DbServerType,
                        Server = x.Server,
                        CompanyDB = x.CompanyDB,
                        DbUserName = x.DbUserName,
                        DbPassword = x.DbPassword,
                        LicenseServer = x.LicenseServer,
                        UserName = x.UserName,
                        Password = x.Password
                    };
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"CONF: {ex.ToString()}");
            }
        }

        public static SAPbobsCOM.Recordset RunQuery(string query)
        {
            try
            {
                oRec = (SAPbobsCOM.Recordset)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRec.DoQuery(query);

                return oRec;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static object Release(object objeto)
        {
            if (objeto != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto);
                Query = null;
                GC.Collect();
            }
            return null;
        }

        public static void Log(string sMensaje)
        {
            string BD = "";

            if (Globals.oCompany != null && Globals.oCompany.Connected)
            {
                BD = Globals.oCompany.CompanyDB;
            }

            WriteTxt(BD, sMensaje, "", "");
        }

        public static void Log(Exception ex)
        {
            // Get stack trace for the exception with source file information
            var trace = new StackTrace(ex, true);

            // Get the top stack frame
            var frame = trace.GetFrame(0);

            WriteTxt("", ex.Message, frame.GetMethod().Name, frame.GetFileLineNumber().ToString());
        }

        public static void WriteTxt(string BD, string x, string Method, string Line)//, string filename
        {
            string path = null;
            path = AppDomain.CurrentDomain.BaseDirectory + "\\LOG";
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }

            string filename = DateTime.Now.ToString("yyyyMMdd");
            string FILE_NAME = path + "\\" + filename + ".txt";
            if (!System.IO.File.Exists(FILE_NAME))
            {
                System.IO.File.Create(FILE_NAME).Dispose();
            }

            string texto = string.Empty;
            if (x.Contains("Iniciado"))
                texto = "***************************************************\n";

            texto += BD + " ";
            texto += (Globals.Right("00" + DateTime.Now.Hour.ToString(), 2) + ":" +
            Globals.Right("00" + DateTime.Now.Minute.ToString(), 2) + ":" +
            Globals.Right("00" + DateTime.Now.Second.ToString(), 2)) + "-(" + Method + ": " + Line + ")-";
            texto += x;

            System.IO.StreamWriter objWriter = new System.IO.StreamWriter(FILE_NAME, true);
            objWriter.WriteLine(texto);
            objWriter.Close();
        }

        public static string Right(string value, int length)
        {
            return value.Substring(value.Length - length);
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
                throw ex;
            }
        }

        public static String RetirarMonedayDecimal(string cadena, int numdec = 0)
        {
            try
            {
                int CantDecimales = 2;

                if (numdec != 0)
                    CantDecimales = numdec;

                Char delimiter = '.';
                String final = cadena;

                foreach (DataRow drMoneda in Globals.dtMonedas.Rows)
                {
                    final = final.Replace(drMoneda["CurrCode"].ToString(), "").Trim();
                }

                final = final.Replace("$", "").Trim();
                final = final.Replace("US", "").Trim();
                final = final.Replace(",", "").Trim();

                if (final.Trim().Length == 0)
                    final = "0.00";
                else
                {
                    final = Math.Round(Convert.ToDouble(final), CantDecimales, MidpointRounding.AwayFromZero).ToString();
                    if (final.ToString().Contains(delimiter))
                    {
                        int decimales = final.Length - (final.LastIndexOf(".") == -1 ? 0 : final.LastIndexOf(".")) - 1;
                        final = final.PadRight(final.Length + CantDecimales - decimales, '0');
                    }
                    else
                        final = final + ".00";
                }

                return final;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void ObtenerConfiguracionInicial()
        {
            try
            {
                SAPDI.ObtenerDatosOADM();
                SAPDI.ObtenerDatosADM1();
                SAPDI.ObtenerMonedas();
                SAPDI.ObtenerConfig();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static System.Data.DataTable RecorsetToDataTable(Recordset SAPRecordset)
        {
            System.Data.DataTable dtTable = new System.Data.DataTable();
            System.Data.DataColumn NewCol;
            DataRow NewRow;
            int ColCount;

            try
            {
                for (ColCount = 0; ColCount < SAPRecordset.Fields.Count; ColCount++)
                {
                    NewCol = new System.Data.DataColumn(SAPRecordset.Fields.Item(ColCount).Name);
                    dtTable.Columns.Add(NewCol);
                }

                while (!(SAPRecordset.EoF))
                {
                    NewRow = dtTable.NewRow();
                    //populate each column in the row we're creating
                    for (ColCount = 0; ColCount < SAPRecordset.Fields.Count; ColCount++)
                        NewRow[SAPRecordset.Fields.Item(ColCount).Name] = SAPRecordset.Fields.Item(ColCount).Value;

                    //Add the row to the datatable
                    dtTable.Rows.Add(NewRow);

                    SAPRecordset.MoveNext();
                }

                return dtTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static String MontoEnLetras(String num, String nombreMoneda)
        {
            String res, dec = "";
            Int64 entero;
            Int32 decimales;
            Double nro;
            try
            {
                nro = Convert.ToDouble(num);
            }
            catch
            {
                return "";
            }
            entero = Convert.ToInt64(Math.Truncate(nro));
            decimales = Convert.ToInt32(Math.Round((nro - entero) * 100, 2));
            //if (decimales > 0)//D.P.M.
            //{
            dec = " CON " + decimales.ToString().PadLeft(2, '0') + "/100";
            //}
            res = ConvertirATexto(Convert.ToDouble(entero)) + dec + " " + nombreMoneda.ToUpper();
            return res;
        }

        private static String ConvertirATexto(Double value)
        {
            try
            {
                String Num2Text = "";
                value = Math.Truncate(value);

                if (value == 0) Num2Text = "CERO";
                else if (value == 1) Num2Text = "UNO";
                else if (value == 2) Num2Text = "DOS";
                else if (value == 3) Num2Text = "TRES";
                else if (value == 4) Num2Text = "CUATRO";
                else if (value == 5) Num2Text = "CINCO";
                else if (value == 6) Num2Text = "SEIS";
                else if (value == 7) Num2Text = "SIETE";
                else if (value == 8) Num2Text = "OCHO";
                else if (value == 9) Num2Text = "NUEVE";
                else if (value == 10) Num2Text = "DIEZ";
                else if (value == 11) Num2Text = "ONCE";
                else if (value == 12) Num2Text = "DOCE";
                else if (value == 13) Num2Text = "TRECE";
                else if (value == 14) Num2Text = "CATORCE";
                else if (value == 15) Num2Text = "QUINCE";
                else if (value < 20) Num2Text = "DIECI" + ConvertirATexto(value - 10);
                else if (value == 20) Num2Text = "VEINTE";
                else if (value < 30) Num2Text = "VEINTI" + ConvertirATexto(value - 20);
                else if (value == 30) Num2Text = "TREINTA";
                else if (value == 40) Num2Text = "CUARENTA";
                else if (value == 50) Num2Text = "CINCUENTA";
                else if (value == 60) Num2Text = "SESENTA";
                else if (value == 70) Num2Text = "SETENTA";
                else if (value == 80) Num2Text = "OCHENTA";
                else if (value == 90) Num2Text = "NOVENTA";
                else if (value < 100) Num2Text = ConvertirATexto(Math.Truncate(value / 10) * 10) + " Y " + ConvertirATexto(value % 10);
                else if (value == 100) Num2Text = "CIEN";
                else if (value < 200) Num2Text = "CIENTO " + ConvertirATexto(value - 100);
                else if ((value == 200) || (value == 300) || (value == 400) || (value == 600) || (value == 800)) Num2Text = ConvertirATexto(Math.Truncate(value / 100)) + "CIENTOS";
                else if (value == 500) Num2Text = "QUINIENTOS";
                else if (value == 700) Num2Text = "SETECIENTOS";
                else if (value == 900) Num2Text = "NOVECIENTOS";
                else if (value < 1000) Num2Text = ConvertirATexto(Math.Truncate(value / 100) * 100) + " " + ConvertirATexto(value % 100);
                else if (value == 1000) Num2Text = "MIL";
                else if (value < 2000) Num2Text = "MIL " + ConvertirATexto(value % 1000);
                else if (value < 1000000) { Num2Text = ConvertirATexto(Math.Truncate(value / 1000)) + " MIL"; if ((value % 1000) > 0) Num2Text = Num2Text + " " + ConvertirATexto(value % 1000); }
                else if (value == 1000000) Num2Text = "UN MILLON";
                else if (value < 2000000) Num2Text = "UN MILLON " + ConvertirATexto(value % 1000000);
                else if (value < 1000000000000) { Num2Text = ConvertirATexto(Math.Truncate(value / 1000000)) + " MILLONES "; if ((value - Math.Truncate(value / 1000000) * 1000000) > 0) Num2Text = Num2Text + " " + ConvertirATexto(value - Math.Truncate(value / 1000000) * 1000000); }
                else if (value == 1000000000000) Num2Text = "UN BILLON";
                else if (value < 2000000000000) Num2Text = "UN BILLON " + ConvertirATexto(value - Math.Truncate(value / 1000000000000) * 1000000000000); else { Num2Text = ConvertirATexto(Math.Truncate(value / 1000000000000)) + " BILLONES"; if ((value - Math.Truncate(value / 1000000000000) * 1000000000000) > 0) Num2Text = Num2Text + " " + ConvertirATexto(value - Math.Truncate(value / 1000000000000) * 1000000000000); }
                return Num2Text;
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

        public static String ObtenerDescripcionMoneda(String MonedaSAP)
        {
            try
            {
                var value = (from DataRow myRow in Globals.dtMonedas.Rows
                             where myRow["ISOCurrCod"].ToString() == MonedaSAP.Trim()
                             select myRow["CurrName"]).First().ToString().ToUpper();

                if (value.ToString().Trim().Length == 0)
                    throw new Exception("No se encuentra definido la descripción de Moneda para " + MonedaSAP + ". Por favor verifique el maestro de monedas.");
                else
                    return value.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static String SumarCadenas(String Cadena1, String Cadena2)
        {
            try
            {
                Double Cantidad1 = Convert.ToDouble(Cadena1);
                Double Cantidad2 = Convert.ToDouble(Cadena2);
                return RetirarMonedayDecimal((Cantidad1 + Cantidad2).ToString());
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void ObtenerAnticiposPaperless(ref Entities.Paperless.Documento oBEiFEDoc, string SAPTable, int iDocEntry)
        {
            try
            {
                Globals.Query = ServicioConectorFE.Properties.Resources.ObtenerAnticipos;
                Globals.Query = string.Format(Globals.Query, SAPTable.Substring(1, 3) + "9", iDocEntry);
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();

                if (Globals.oRec.RecordCount == 0)
                {
                    throw new Exception("No se ha logrado obtener datos del anticipo. Por favor verifique.");
                }
                else
                {

                    oBEiFEDoc.ListAnticipo = new List<Entities.Paperless.Anticipo>();
                    Globals.oRec.MoveFirst();
                    for (int i = 0; i < Globals.oRec.RecordCount; i++)
                    {
                        Entities.Paperless.Anticipo Anticipo = new Entities.Paperless.Anticipo();
                        Anticipo.rutEmisorDocumento = oBEiFEDoc.emisor.ruc;
                        Anticipo.tipoDocumento = Globals.oRec.Fields.Item(0).Value.ToString();
                        Anticipo.serieDocumento = Globals.oRec.Fields.Item(1).Value.ToString();
                        Anticipo.correlativoDocumento = Globals.oRec.Fields.Item(2).Value.ToString();
                        Anticipo.importeDocumento = Globals.RetirarMonedayDecimal(Convert.ToString(Globals.oRec.Fields.Item(3).Value));
                        Anticipo.totalAnticipos = Globals.RetirarMonedayDecimal(Convert.ToString(Globals.oRec.Fields.Item(4).Value));
                        Anticipo.valorIGV = Globals.RetirarMonedayDecimal(Convert.ToString(Globals.oRec.Fields.Item(5).Value));
                        Anticipo.fechaDocumento = Globals.oRec.Fields.Item(6).Value.ToString();
                        oBEiFEDoc.ListAnticipo.Add(Anticipo);
                        Globals.oRec.MoveNext();
                    }
                }
                Globals.Release(Globals.oRec);
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


        public static void ObtenerAnticiposEstela(ref Entities.Estela.Documento oBEiFEDoc, string SAPTable, int iDocEntry)
        {
            try
            {
                Globals.Query = ServicioConectorFE.Properties.Resources.ObtenerAnticipos;
                Globals.Query = string.Format(Globals.Query, SAPTable.Substring(1, 3) + "9", iDocEntry);
                Globals.RunQuery(Globals.Query);
                Globals.oRec.MoveFirst();

                if (Globals.oRec.RecordCount == 0)
                {
                    throw new Exception("No se ha logrado obtener datos del anticipo. Por favor verifique.");
                }
                else
                {

                    oBEiFEDoc.ListAnticipo = new List<Anticipo>();
                    Globals.oRec.MoveFirst();
                    for (int i = 0; i < Globals.oRec.RecordCount; i++)
                    {
                        Anticipo Anticipo = new Anticipo();
                        Anticipo.rutEmisorDocumento = oBEiFEDoc.emisor.ruc;
                        Anticipo.tipoDocumento = Globals.oRec.Fields.Item(0).Value.ToString();
                        Anticipo.serieDocumento = Globals.oRec.Fields.Item(1).Value.ToString();
                        Anticipo.correlativoDocumento = Globals.oRec.Fields.Item(2).Value.ToString();
                        Anticipo.importeDocumento = Globals.RetirarMonedayDecimal(Convert.ToString(Globals.oRec.Fields.Item(3).Value));
                        Anticipo.totalAnticipos = Globals.RetirarMonedayDecimal(Convert.ToString(Globals.oRec.Fields.Item(4).Value));
                        Anticipo.valorIGV = Globals.RetirarMonedayDecimal(Convert.ToString(Globals.oRec.Fields.Item(5).Value));
                        Anticipo.fechaDocumento = Globals.oRec.Fields.Item(6).Value.ToString();
                        oBEiFEDoc.ListAnticipo.Add(Anticipo);
                        Globals.oRec.MoveNext();
                    }
                }
                Globals.Release(Globals.oRec);
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

        public static void ValidaCamposCRE(Entities.Estela.DocumentoRetencionFE oBeFEDOC)
        {
            try
            {
                if (string.IsNullOrEmpty(oBeFEDOC.regimenRetencion))
                    throw new Exception("Debe seleccionar el régimen de retención");
            }
            catch (Exception ex)
            {
                throw;
            }
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

        public static void ValidaCamposBaja(Entities.SAP.Baja oBaja)
        {
            try
            {
                if (string.IsNullOrEmpty(oBaja.notaSustento))
                    throw new Exception("Debe indicar un motivo de baja");
            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
}
