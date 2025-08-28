using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServicioConectorFE.Entities.Estela
{
    public class Documento
    {  //EN
        public string DocEntry { get; set; }
        public string documentoExportacion { get; set; }
        public bool TransGratuita { get; set; }
        public string CardCode { get; set; }
        public string PayToCode { get; set; }
        public string ShipToCode { get; set; }
        public string DocType { get; set; }
        public string tipoComprobante { get; set; }
        public string serie { get; set; }
        public string correlativo { get; set; }
        public string notaTipo { get; set; }
        public string notaSustento { get; set; }
        public string fechaEmision { get; set; }
        public string tipoMoneda { get; set; }
        public string TipoCambio { get; set; }

        //ENEX
        public string tipoOperacion { get; set; }
        public string ordenCompra { get; set; }
        public string redondeo { get; set; }
        public string valorTotalAnticipos { get; set; }
        public string fechaVencimiento { get; set; }
        public string horaEmision { get; set; }
        public string codigoEstablecimiento { get; set; }
        public string pedido { get; set; }
        public string guias { get; set; }
        public string importeTotal { get; set; }
        public string varlorDescuentoGlobal { get; set; }

        public string vendedor { get; set; }
        public string credito { get; set; }

        //Detraccion
        public string TieneDetraccion { get; set; }
        public string codigoDetraccion { get; set; }
        public string totalDetraccion { get; set; }
        public string totalDetraccionMD { get; set; }
        public string porcentajeDetraccion { get; set; }

        public string condicionPago { get; set; }
        public string condicionPagoDesc { get; set; }
        public string observaciones { get; set; }
        public string Responsable { get; set; }

        //Montos
        public string valorOperacionesGravadas { get; set; }
        public string valorOperacionesExoneradas { get; set; }
        public string valorOperacionesInafectas { get; set; }
        public string valorOperacionesGratuitas { get; set; }
        public string valorDescuentoGlobal { get; set; }
        public string valorTotalIGVAnticipos { get; set; }
        public string otrosCargos { get; set; }

        public string tasaIgv { get; set; }
        public string valorIgv { get; set; }

        //Campos Personalizados
        public string U_EXT_NAVE { get; set; }
        public string U_EXT_MUELLE { get; set; }
        public string U_EXT_CALLNUMBER { get; set; }
        public string U_EXT_LINEANAV { get; set; }
        public string U_EXT_RESPSOLIDA { get; set; }
        public string U_EXT_OSNRO { get; set; }
        public string U_EXT_FECARRIBO { get; set; }
        public string Comments { get; set; }

        //Desglosado por tipo dato
        public Emisor emisor { get; set; }
        public Adquiriente adquiriente { get; set; }
        public List<Anticipo> ListAnticipo { get; set; }
        public List<Referencia> ListReferencia { get; set; }
        public List<DocLine> Detalles { get; set; }
    }

    public class Emisor
    {
        public string ruc { get; set; }
        public string tipoIdentificacion { get; set; }
        public string nombreComercial { get; set; }
        public string razonSocial { get; set; }
        public string ubigeo { get; set; }
        public string direccion { get; set; }
        public string departamento { get; set; }
        public string provincia { get; set; }
        public string distrito { get; set; }
    }

    public class Adquiriente
    {
        public string numeroIdentificacion { get; set; }
        public string tipoIdentificacion { get; set; }
        public string nombre { get; set; }
        public string ubigeo { get; set; }
        public string direccionFiscal { get; set; }
        public string codigoPais { get; set; }

        public string telefono { get; set; }
        public string fax { get; set; }
        public string email { get; set; }
        public string obra { get; set; }
    }

    public class Referencia
    {
        public string tipoDocumento { get; set; }
        public string serie { get; set; }
        public string correlativo { get; set; }
        public string fechaEmision { get; set; }
    }

    public class Anticipo
    {
        public string fechaDocumento { get; set; }
        public string rutEmisorDocumento { get; set; }
        public string tipoDocumento { get; set; }
        public string serieDocumento { get; set; }
        public string correlativoDocumento { get; set; }
        public string importeDocumento { get; set; }
        public string totalAnticipos { get; set; }
        public string valorIGV { get; set; }
    }

    public class DocLine
    {
        public string BaseEntry { get; set; }
        public string posicion { get; set; }
        public string codigo { get; set; }
        public string descripcion { get; set; }
        public string cantidadUnidades { get; set; }
        public string unidadMedida { get; set; }
        public string unidadMedidaSAP { get; set; }
        public string valorUnitario { get; set; }
        public string precioUnitario { get; set; }
        public string IgvTipo { get; set; }
        public string IgvMonto { get; set; }
        public string IgvMontoGratuito { get; set; }
        public string descuentoPct { get; set; }
        public string descuento { get; set; }
        public string valorItem { get; set; }
        public string valorItemGratutito { get; set; }
        public string cobroMinCon { get; set; }
        public string CodigoImpuesto { get; set; }
        public double tasaIGV { get; set; }
        public string Rate { get; set; }
        public string ValorCodigo { get; set; }
        public string CodigoTributo { get; set; }
        public string CodigoTributoInt { get; set; }
        public string NombreTributo { get; set; }
        public string OperacionAfecta { get; set; }
        public string TaxOnly { get; set; }
        public string Glosa { get; set; }
    }
}