using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServicioConectorFE.Core
{
    public class Structure
    {
        public static void CrearMetaData()
        {
            List<List<String>> TrueFalseValues = new List<List<String>>();
            agregarValorValido("Y", "SI", ref TrueFalseValues);
            agregarValorValido("N", "NO", ref TrueFalseValues);

            List<List<String>> TipoEtiquetaValues = new List<List<String>>();
            agregarValorValido("C", "Cabecera (Observacion Adicional)", ref TipoEtiquetaValues);
            agregarValorValido("D", "Detalle (Columna Adicional)", ref TipoEtiquetaValues);

            crearCampo("OCTG", "EXX_ESCRED", "Es crédito?", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 2, "N", "", false, TrueFalseValues);

            crearCampo("NNM1", "EXX_ANEXO", "Ruta a Anexos", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254, "", "", false, null);
        }

        private static void agregarValorValido(String Value, String Description, ref List<List<String>> ValidValues)
        {
            List<String> ValidValue = new List<String>();
            ValidValue.Add(Value);
            ValidValue.Add(Description);
            ValidValues.Add(ValidValue);
        }

        private static bool crearTabla(string tabla, string nombretabla, SAPbobsCOM.BoUTBTableType tipo = SAPbobsCOM.BoUTBTableType.bott_NoObject)
        {
            SAPbobsCOM.UserTablesMD oTablaUser = (SAPbobsCOM.UserTablesMD)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
            try
            {
                if (!oTablaUser.GetByKey(tabla))
                {
                    oTablaUser.TableName = tabla;
                    oTablaUser.TableDescription = nombretabla;
                    oTablaUser.TableType = tipo;

                    int RetVal = oTablaUser.Add();
                    if ((RetVal != 0))
                    {
                        String errMsg;
                        int errCode;
                        Globals.oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception(errMsg);
                    }
                    else
                        return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTablaUser);
            }
        }

        private static void crearCampo(string tabla, string campo, string descripcion, SAPbobsCOM.BoFieldTypes tipo,
            SAPbobsCOM.BoFldSubTypes subtipo, int tamaño, string ValorPorDefecto, string sLinkedTable,
            Boolean Mandatory, List<List<String>> ValidValues)//CHG
        {
            int existeCampo = 0;

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string cadena = "select \"FieldID\" from CUFD where (\"TableID\"='" + tabla + "' or \"TableID\"='@" + tabla + "') and \"AliasID\"='" + campo + "'";
            rs.DoQuery(cadena);

            existeCampo = rs.RecordCount;
            int FieldID = Convert.ToInt32(rs.Fields.Item(0).Value);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            rs = null;

            SAPbobsCOM.UserFieldsMD oCampo = (SAPbobsCOM.UserFieldsMD)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            try
            {
                if (existeCampo == 0)//Crear
                {
                    oCampo.TableName = tabla;
                    oCampo.Name = campo;
                    oCampo.Description = descripcion;
                    oCampo.Type = tipo;
                    oCampo.SubType = subtipo;
                    oCampo.Mandatory = Mandatory ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;

                    if (tamaño > 0)
                    {
                        oCampo.EditSize = tamaño;
                    }

                    if (sLinkedTable.ToString() != "")
                        oCampo.LinkedTable = sLinkedTable;

                    if (ValidValues != null)
                    {
                        foreach (List<String> ValidValue in ValidValues)
                        {
                            oCampo.ValidValues.Value = ValidValue[0];
                            oCampo.ValidValues.Description = ValidValue[1];
                            oCampo.ValidValues.Add();
                        }
                    }

                    if (ValorPorDefecto.ToString() != "")
                    {
                        oCampo.DefaultValue = ValorPorDefecto;
                    }

                    int RetVal = oCampo.Add();
                    if (RetVal != 0)
                    {
                        String errMsg;
                        int errCode;
                        Globals.oCompany.GetLastError(out errCode, out errMsg);
                        throw new Exception(errMsg);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCampo);
            }
        }
    }
}
