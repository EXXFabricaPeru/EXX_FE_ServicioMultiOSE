using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServicioConectorFE.Core
{
    public class Setup
    {
        public static bool ValidaVersion(string addOnName, String addOnVersion)
        {
            bool retorno = false;
            try
            {
                if (!checkCampoBD("@EXX_SETUP", "U_EXX_VERS"))
                {
                    creaTablaMD("EXX_SETUP", "Setup de AddOns de EXX", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                    //creaCampoMD("EXX_SETUP", "EXX_ADDN", "Nombre del AddOn", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
                    //creaCampoMD("EXX_SETUP", "EXX_VERS", "Version del AddOn", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 100);
                    //creaCampoMD("EXX_SETUP", "EXX_RUTA", "Ruta auxiliar para AddOn", SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, 254);
                    //confirmarVersion(addOnName, addOnVersion);
                }
            }
            catch (Exception ex)
            {
                Globals.Log(ex.Message);
                retorno = false;
            }
            return retorno;
        }

        public static bool checkCampoBD(string Tabla, string Campo)
        {
            bool retorno = false;
            try
            {
                string strSQLBD = null;
                SAPbobsCOM.Recordset oLocalBD = default(SAPbobsCOM.Recordset);
                oLocalBD = (SAPbobsCOM.Recordset)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (Globals.IsHana() == true)
                {
                    strSQLBD = "SELECT COLUMN_NAME FROM SYS.M_CS_COLUMNS WHERE COLUMN_NAME = '" + Campo + "' AND TABLE_NAME = '" + Tabla + "' AND SCHEMA_NAME = '" + Globals.oCompany.CompanyDB.ToString() + "'";
                    oLocalBD.DoQuery(strSQLBD);
                }
                else
                {
                    strSQLBD = "SELECT column_name ";
                    strSQLBD += "FROM [" + Globals.oCompany.CompanyDB + "].INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME = '" + Campo + "' AND Table_Name ='" + Tabla + "'";
                    oLocalBD.DoQuery(strSQLBD);
                }
                if (oLocalBD.EoF == false)
                {
                    retorno = true;
                }
                Globals.Release(oLocalBD);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return retorno;
        }
        public static void creaTablaMD(string NbTabla, string DescTabla, SAPbobsCOM.BoUTBTableType TablaTipo)
        {
            SAPbobsCOM.UserTablesMD oUserTablesMD = default(SAPbobsCOM.UserTablesMD);
            try
            {
                int iVer = 0;
                oUserTablesMD = (SAPbobsCOM.UserTablesMD)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);

                if (!oUserTablesMD.GetByKey(NbTabla))
                {
                    SAPbobsCOM.UserTablesMD tablaACrear = (SAPbobsCOM.UserTablesMD)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                    tablaACrear.TableName = string.Format(NbTabla);
                    tablaACrear.TableDescription = string.Format(DescTabla);
                    tablaACrear.TableType = TablaTipo;

                    int retX = 0;
                    string strSQLx = "";
                    retX = tablaACrear.Add();
                    if (!(retX == 0))
                    {
                        iVer = iVer + 1;
                        Globals.oCompany.GetLastError(out retX, out strSQLx);
                    }
                    else
                    {
                        Globals.Log(Globals.oCompany.CompanyDB + ": tabla " + NbTabla + " creada con éxito");
                    }
                    Globals.Release(tablaACrear);
                }
                Globals.Release(oUserTablesMD);
            }
            catch (Exception ex)
            {
                Globals.Log(ex);
            }
        }
        //public static void creaCampoMD(string NbTabla, string NbCampo, string DescCampo, SAPbobsCOM.BoFieldTypes TipoDato, SAPbobsCOM.BoFldSubTypes subtipo = SAPbobsCOM.BoFldSubTypes.st_None, int Tamaño = 10, SAPbobsCOM.BoYesNoEnum Obligatorio = SAPbobsCOM.BoYesNoEnum.tNO, string[] validValues = null, string[] validDescription = null, string valorPorDef = "", string tablaVinculada = "")
        //{
        //    SAPbobsCOM.UserFieldsMD oUserFieldsMD = default(SAPbobsCOM.UserFieldsMD);
        //    try
        //    {
        //        oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

        //        oUserFieldsMD.TableName = NbTabla;
        //        oUserFieldsMD.Name = NbCampo;
        //        oUserFieldsMD.Description = DescCampo;
        //        oUserFieldsMD.Type = TipoDato;
        //        if (TipoDato != SAPbobsCOM.BoFieldTypes.db_Date)
        //            oUserFieldsMD.EditSize = Tamaño;
        //        if (TipoDato == SAPbobsCOM.BoFieldTypes.db_Float)
        //            oUserFieldsMD.SubType = subtipo;

        //        if (!string.IsNullOrEmpty(tablaVinculada))
        //        {
        //            oUserFieldsMD.LinkedTable = tablaVinculada;
        //        }
        //        else
        //        {
        //            if ((validValues != null))
        //            {
        //                for (int i = 0; i <= validValues.Length - 1; i++)
        //                {
        //                    if (validDescription == null)
        //                    {
        //                        oUserFieldsMD.ValidValues.Description = validValues[i];
        //                    }
        //                    else
        //                    {
        //                        oUserFieldsMD.ValidValues.Description = validDescription[i];
        //                    }
        //                    oUserFieldsMD.ValidValues.Value = validValues[i];
        //                    oUserFieldsMD.ValidValues.Add();
        //                }
        //            }

        //            if (!string.IsNullOrEmpty(valorPorDef))
        //            {
        //                oUserFieldsMD.DefaultValue = valorPorDef;
        //                oUserFieldsMD.Mandatory = Obligatorio;
        //            }
        //        }

        //        int retX = 0;
        //        string strSQLx = "";
        //        retX = oUserFieldsMD.Add();

        //        if (retX != 0)
        //        {
        //            Globals.oCompany.GetLastError(out retX, out strSQLx);
        //        }
        //        else
        //        {
        //            Globals.SBO_Application.StatusBar.SetText("Campo " + NbCampo + " " + DescCampo + ": Creado con éxito", SAPbouiCOM.BoMessageTime.bmt_Short, (SAPbouiCOM.BoStatusBarMessageType)SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        //        }
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
        //        oUserFieldsMD = null;
        //        GC.Collect();
        //        return;
        //    }
        //    catch (Exception ex)
        //    {
        //        Globals.SBO_Application.MessageBox(ex.Message);
        //    }
        //}
        //public string getCorrelativo(string CampoMax, string Tabla, string condicion = "", int primerCorrelativo = 1)
        //{
        //    SAPbobsCOM.Recordset oMax = (SAPbobsCOM.Recordset)Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    string Srt = primerCorrelativo.ToString();
        //    try
        //    {
        //        poner aca un ifhana para HanaCorrInstall
        //        if (Globals.IsHana() == true)
        //            Srt = AddonEAR.Properties.Resources.HanaCorrInstall;
        //        if (Globals.IsHana() == false)
        //            Srt = AddonEAR.Properties.Resources.SQLCorrInstall;
        //        if (!string.IsNullOrEmpty(condicion))
        //        {
        //            int numero = primerCorrelativo - 1;
        //            poner acá otro ifhana para ese Srt
        //            if (Globals.IsHana() == false)
        //                Srt = "SELECT ISNULL(MAX(CAST(" + CampoMax + " AS numeric)), " + numero + ") + 1 AS Numero FROM (SELECT * FROM OWHS WHERE " + condicion + ") AS X WHERE " + condicion;
        //            if (Globals.IsHana() == true)
        //                Srt = "SELECT ISNULL(MAX(CAST(" + CampoMax + " AS numeric)), " + numero + ") + 1 AS Numero FROM (SELECT * FROM OWHS WHERE " + condicion + ") AS X WHERE " + condicion;
        //        }
        //        oMax.DoQuery(Srt);
        //        Srt = (oMax.EoF == true ? primerCorrelativo.ToString() : oMax.Fields.Item("Numero").Value.ToString());
        //    }
        //    catch (Exception ex)
        //    {
        //        Globals.SBO_Application.MessageBox(ex.Message);
        //        Srt = "0";
        //    }
        //    finally
        //    {
        //        Globals.Release(oMax);
        //    }
        //    return Srt;
        //}

        //private int CompareVersion(string a, string b)
        //{
        //    int a1, b1;
        //    string aa = a.Replace(".", "");
        //    string bb = b.Replace(".", "");
        //    if (aa.Length > bb.Length)
        //    {
        //        bb = bb.PadRight(bb.Length + (aa.Length - bb.Length), '0');
        //    }
        //    else if (aa.Length < bb.Length)
        //    {
        //        aa = aa.PadRight(aa.Length + (bb.Length - aa.Length), '0');
        //    }
        //    a1 = Convert.ToInt16(aa);
        //    b1 = Convert.ToInt16(bb);
        //    if (a1 == b1) return 0;
        //    else if (a1 < b1) return 1;
        //    else return 2;
        //}
    }
}
