using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServicioConectorFE.Entities.Nubefact
{
    public class Documento
    {

        [JsonProperty("operacion")]
        public string Operacion { get; set; }

        [JsonProperty("tipo_de_comprobante")]
        public string TipoDeComprobante { get; set; }

        [JsonProperty("serie")]
        public string Serie { get; set; }

        [JsonProperty("numero")]
        public int Numero { get; set; }

        [JsonProperty("sunat_transaction")]
        public int? SunatTransaction { get; set; }

        [JsonProperty("cliente_tipo_de_documento")]
        public int? ClienteTipoDeDocumento { get; set; }

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

        [JsonProperty("fecha_de_vencimiento")]
        public string FechaDeVencimiento { get; set; }

        [JsonProperty("moneda")]
        public string Moneda { get; set; }

        [JsonProperty("tipo_de_cambio")]
        public string TipoDeCambio { get; set; }

        [JsonProperty("porcentaje_de_igv")]
        public string PorcentajeDeIgv { get; set; }

        [JsonProperty("descuento_global")]
        public string DescuentoGlobal { get; set; }

        [JsonProperty("total_descuento")]
        public string TotalDescuento { get; set; }

        [JsonProperty("total_anticipo")]
        public string TotalAnticipo { get; set; }

        [JsonProperty("total_gravada")]
        public string TotalGravada { get; set; }

        [JsonProperty("total_inafecta")]
        public string TotalInafecta { get; set; }

        [JsonProperty("total_exonerada")]
        public string TotalExonerada { get; set; }

        [JsonProperty("total_igv")]
        public string TotalIgv { get; set; }

        [JsonProperty("total_gratuita")]
        public string TotalGratuita { get; set; }

        [JsonProperty("total_otros_cargos")]
        public string TotalOtrosCargos { get; set; }

        [JsonProperty("total")]
        public string Total { get; set; }

        [JsonProperty("percepcion_tipo")]
        public string PercepcionTipo { get; set; }

        [JsonProperty("percepcion_base_imponible")]
        public string PercepcionBaseImponible { get; set; }

        [JsonProperty("total_percepcion")]
        public string TotalPercepcion { get; set; }

        [JsonProperty("total_incluido_percepcion")]
        public string TotalIncluidoPercepcion { get; set; }

        [JsonProperty("detraccion")]
        public string Detraccion { get; set; }
        [JsonProperty("detraccion_tipo")]
        public string detracciontipo { get; set; }
        [JsonProperty("detraccion_total")]
        public string detracciontotal { get; set; }

        [JsonProperty("observaciones")]
        public string Observaciones { get; set; }

        [JsonProperty("documento_que_se_modifica_tipo")]
        public string DocumentoQueSeModificaTipo { get; set; }

        [JsonProperty("documento_que_se_modifica_serie")]
        public string DocumentoQueSeModificaSerie { get; set; }

        [JsonProperty("documento_que_se_modifica_numero")]
        public string DocumentoQueSeModificaNumero { get; set; }

        [JsonProperty("tipo_de_nota_de_credito")]
        public string TipoDeNotaDeCredito { get; set; }

        [JsonProperty("tipo_de_nota_de_debito")]
        public string TipoDeNotaDeDebito { get; set; }

        [JsonProperty("enviar_automaticamente_a_la_sunat")]
        public string EnviarAutomaticamenteALaSunat { get; set; }

        [JsonProperty("enviar_automaticamente_al_cliente")]
        public string EnviarAutomaticamenteAlCliente { get; set; }

        [JsonProperty("codigo_unico")]
        public string CodigoUnico { get; set; }

        [JsonProperty("condiciones_de_pago")]
        public string CondicionesDePago { get; set; }

        [JsonProperty("medio_de_pago")]
        public string MedioDePago { get; set; }

        [JsonProperty("placa_vehiculo")]
        public string PlacaVehiculo { get; set; }

        [JsonProperty("orden_compra_servicio")]
        public string OrdenCompraServicio { get; set; }

        [JsonProperty("tabla_personalizada_codigo")]
        public string TablaPersonalizadaCodigo { get; set; }

        [JsonProperty("formato_de_pdf")]
        public string FormatoDePdf { get; set; }

        [JsonProperty("items")]
        public List<Item> Items { get; set; }
    }

    public class Item
    {
        [JsonProperty("unidad_de_medida")]
        public string UnidadDeMedida { get; set; }

        [JsonProperty("codigo")]
        public string Codigo { get; set; }

        [JsonProperty("descripcion")]
        public string Descripcion { get; set; }

        [JsonProperty("cantidad")]
        public string Cantidad { get; set; }

        [JsonProperty("valor_unitario")]
        public string ValorUnitario { get; set; }

        [JsonProperty("precio_unitario")]
        public string PrecioUnitario { get; set; }

        [JsonProperty("descuento")]
        public string Descuento { get; set; }

        [JsonProperty("subtotal")]
        public string Subtotal { get; set; }

        [JsonProperty("tipo_de_igv")]
        public string TipoDeIgv { get; set; }

        [JsonProperty("igv")]
        public string Igv { get; set; }

        [JsonProperty("total")]
        public string Total { get; set; }

        [JsonProperty("anticipo_regularizacion")]
        public string AnticipoRegularizacion { get; set; }

        [JsonProperty("anticipo_documento_serie")]
        public string AnticipoDocumentoSerie { get; set; }

        [JsonProperty("anticipo_documento_numero")]
        public string AnticipoDocumentoNumero { get; set; }

        [JsonProperty("codigo_producto_sunat")]
        public string CodigoProductoSunat { get; set; }
    }

    public class Consulta
    {
        public string operacion { get; set; }
        public string tipo_de_comprobante { get; set; }
        public string serie { get; set; }
        public string numero { get; set; }
    }

    public class Respuesta
    {
        public string errors { get; set; }
        public int tipo { get; set; }
        public int tipo_de_comprobante { get; set; }
        public string serie { get; set; }
        public int numero { get; set; }
        public string url { get; set; }
        public string enlace { get; set; }
        public bool aceptada_por_sunat { get; set; }
        public string sunat_description { get; set; }
        public string sunat_note { get; set; }
        public string sunat_responsecode { get; set; }
        public string sunat_soap_error { get; set; }
        public bool anulado { get; set; }
        public string pdf_zip_base64 { get; set; }
        public string xml_zip_base64 { get; set; }
        public string cdr_zip_base64 { get; set; }
        public string cadena_para_codigo_qr { get; set; }
        public string codigo_hash { get; set; }
        public string codigo_de_barras { get; set; }
        public string key { get; set; }
        public string digest_value { get; set; }
        public string enlace_del_pdf { get; set; }
        public string enlace_del_xml { get; set; }
        public string enlace_del_cdr { get; set; }
    }
}
