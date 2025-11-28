using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServicioConectorFE.Entities.Nubefact
{
    public class Retencion
    {
        [JsonProperty("operacion")]
        public string Operacion { get; set; }

        [JsonProperty("serie")]
        public string Serie { get; set; }

        [JsonProperty("numero")]
        public int Numero { get; set; }

        [JsonProperty("cliente_tipo_de_documento")]
        public int ClienteTipoDeDocumento { get; set; }

        [JsonProperty("cliente_numero_de_documento")]
        public string ClienteNumeroDeDocumento { get; set; }

        [JsonProperty("cliente_denominacion")]
        public string ClienteDenominacion { get; set; }

        [JsonProperty("cliente_direccion")]
        public string ClienteDireccion { get; set; }

        [JsonProperty("cliente_email")]
        public string ClienteEmail { get; set; }

        [JsonProperty("cliente_email_1")]
        public string ClienteEmail1 { get; set; }

        [JsonProperty("cliente_email_2")]
        public string ClienteEmail2 { get; set; }

        [JsonProperty("fecha_de_emision")]
        public string FechaDeEmision { get; set; }

        [JsonProperty("moneda")]
        public string Moneda { get; set; }

        [JsonProperty("tipo_de_tasa_de_retencion")]
        public string TipoDeTasaDeRetencion { get; set; }

        [JsonProperty("total_retenido")]
        public string TotalRetenido { get; set; }

        [JsonProperty("total_pagado")]
        public string TotalPagado { get; set; }

        [JsonProperty("observaciones")]
        public string Observaciones { get; set; }

        [JsonProperty("enviar_automaticamente_a_la_sunat")]
        public string EnviarAutomaticamenteALaSunat { get; set; }

        [JsonProperty("enviar_automaticamente_al_cliente")]
        public string EnviarAutomaticamenteAlCliente { get; set; }

        [JsonProperty("codigo_unico")]
        public string CodigoUnico { get; set; }

        [JsonProperty("formato_de_pdf")]
        public string FormatoDePdf { get; set; }

        [JsonProperty("items")]
        public List<RetencionItem> Items { get; set; }
    }

    public class RetencionItem
    {
        [JsonProperty("documento_relacionado_tipo")]
        public string DocumentoRelacionadoTipo { get; set; }

        [JsonProperty("documento_relacionado_serie")]
        public string DocumentoRelacionadoSerie { get; set; }

        [JsonProperty("documento_relacionado_numero")]
        public int DocumentoRelacionadoNumero { get; set; }

        [JsonProperty("documento_relacionado_fecha_de_emision")]
        public string DocumentoRelacionadoFechaDeEmision { get; set; }

        [JsonProperty("documento_relacionado_moneda")]
        public string DocumentoRelacionadoMoneda { get; set; }

        [JsonProperty("documento_relacionado_total")]
        public string DocumentoRelacionadoTotal { get; set; }

        [JsonProperty("pago_fecha")]
        public string PagoFecha { get; set; }

        [JsonProperty("pago_numero")]
        public string PagoNumero { get; set; }

        [JsonProperty("pago_total_sin_retencion")]
        public string PagoTotalSinRetencion { get; set; }

        [JsonProperty("tipo_de_cambio")]
        public string TipoDeCambio { get; set; }

        [JsonProperty("tipo_de_cambio_fecha")]
        public string TipoDeCambioFecha { get; set; }

        [JsonProperty("importe_retenido")]
        public string ImporteRetenido { get; set; }

        [JsonProperty("importe_retenido_fecha")]
        public string ImporteRetenidoFecha { get; set; }

        [JsonProperty("importe_pagado_con_retencion")]
        public string ImportePagadoConRetencion { get; set; }
    }
}
