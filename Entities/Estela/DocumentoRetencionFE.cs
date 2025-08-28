using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ServicioConectorFE.Entities.Estela
{
    public class DocumentoRetencionFE
    {
        [XmlElement("serieNumeroRetencion")]
        public string serieNumeroRetencion { get; set; }

        [XmlElement("fechaEmision")]
        public string fechaEmision { get; set; }

        [XmlElement("tipoDocumento")]
        public string tipoDocumento { get; set; }

        [XmlElement("numeroDocumentoEmisor")]
        public string numeroDocumentoEmisor { get; set; }

        [XmlElement("tipoDocumentoEmisor")]
        public string tipoDocumentoEmisor { get; set; }

        [XmlElement("correoEmisor")]
        public string correoEmisor { get; set; }

        [XmlElement("correoAdquiriente")]
        public string correoAdquiriente { get; set; }

        [XmlElement("razonSocialEmisor")]
        public string razonSocialEmisor { get; set; }

        [XmlElement("nombreComercialEmisor")]
        public string nombreComercialEmisor { get; set; }

        [XmlElement("razonSocialProveedor")]
        public string razonSocialProveedor { get; set; }

        [XmlElement("ubigeoEmisor")]
        public string ubigeoEmisor { get; set; }

        [XmlElement("direccionEmisor")]
        public string direccionEmisor { get; set; }

        [XmlElement("urbanizacionEmisor")]
        public string urbanizacionEmisor { get; set; }

        [XmlElement("provinciaEmisor")]
        public string provinciaEmisor { get; set; }

        [XmlElement("departamentoEmisor")]
        public string departamentoEmisor { get; set; }

        [XmlElement("distritoEmisor")]
        public string distritoEmisor { get; set; }

        [XmlElement("codigoPaisEmisor")]
        public string codigoPaisEmisor { get; set; }

        [XmlElement("numeroDocumentoProveedor")]
        public string numeroDocumentoProveedor { get; set; }

        [XmlElement("tipoDocumentoProveedor")]
        public string tipoDocumentoProveedor { get; set; }

        [XmlElement("nombreComercialProveedor")]
        public string nombreComercialProveedor { get; set; }

        [XmlElement("ubigeoProveedor")]
        public string ubigeoProveedor { get; set; }

        [XmlElement("direccionProveedor")]
        public string direccionProveedor { get; set; }

        [XmlElement("urbanizacionProveedor")]
        public string urbanizacionProveedor { get; set; }

        [XmlElement("provinciaProveedor")]
        public string provinciaProveedor { get; set; }

        [XmlElement("departamentoProveedor")]
        public string departamentoProveedor { get; set; }

        [XmlElement("distritoProveedor")]
        public string distritoProveedor { get; set; }

        [XmlElement("codigoPaisProveedor")]
        public string codigoPaisProveedor { get; set; }

        [XmlElement("regimenRetencion")]
        public string regimenRetencion { get; set; }

        [XmlElement("tasaRetencion")]
        public string tasaRetencion { get; set; }

        [XmlElement("observaciones")]
        public string observaciones { get; set; }

        [XmlElement("importeTotalRetenido")]
        public string importeTotalRetenido { get; set; }

        [XmlElement("tipoMonedaTotalRetenido")]
        public string tipoMonedaTotalRetenido { get; set; }

        [XmlElement("importeTotalPagado")]
        public string importeTotalPagado { get; set; }

        [XmlElement("tipoMonedaTotalPagado")]
        public string tipoMonedaTotalPagado { get; set; }

        [XmlElement("RetencionItem")]
        public List<RetencionItem> RetencionItems { get; set; }
    }

    public class RetencionItem
    {
        [XmlElement("numeroOrdenItem")]
        public int numeroOrdenItem { get; set; }

        [XmlElement("numeroDocumentoRelacionado")]
        public string numeroDocumentoRelacionado { get; set; }

        [XmlElement("fechaEmisionDocumentoRelacionado")]
        public string fechaEmisionDocumentoRelacionado { get; set; }

        [XmlElement("tipoDocumentoRelacionado")]
        public string tipoDocumentoRelacionado { get; set; }

        [XmlElement("importeTotalDocumentoRelacionado")]
        public string importeTotalDocumentoRelacionado { get; set; }

        [XmlElement("tipoMonedaDocumentoRelacionado")]
        public string tipoMonedaDocumentoRelacionado { get; set; }

        [XmlElement("fechaPago")]
        public string fechaPago { get; set; }

        [XmlElement("numeroPago")]
        public string numeroPago { get; set; }

        [XmlElement("importePagoSinRetencion")]
        public string importePagoSinRetencion { get; set; }

        [XmlElement("monedaPago")]
        public string monedaPago { get; set; }

        [XmlElement("importeRetenido")]
        public string importeRetenido { get; set; }

        [XmlElement("monedaImporteRetenido")]
        public string monedaImporteRetenido { get; set; }

        [XmlElement("fechaRetencion")]
        public string fechaRetencion { get; set; }

        [XmlElement("importeTotalPagarNeto")]
        public string importeTotalPagarNeto { get; set; }

        [XmlElement("monedaMontoNetoPagado")]
        public string monedaMontoNetoPagado { get; set; }

        [XmlElement("monedaReferenciaTipoCambio")]
        public string monedaReferenciaTipoCambio { get; set; }

        [XmlElement("monedaObjetivoTasaCambio")]
        public string monedaObjetivoTasaCambio { get; set; }

        [XmlElement("factorTipoCambioMoneda")]
        public string factorTipoCambioMoneda { get; set; }

        [XmlElement("fechaCambio")]
        public string fechaCambio { get; set; }
    }
}
