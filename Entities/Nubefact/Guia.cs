using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServicioConectorFE.Entities.Nubefact
{
    public class Guia
    {
        [JsonProperty("operacion")]
        public string Operacion { get; set; }

        [JsonProperty("tipo_de_comprobante")]
        public int TipoDeComprobante { get; set; }

        [JsonProperty("serie")]
        public string Serie { get; set; }

        [JsonProperty("numero")]
        public string Numero { get; set; }

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

        // Mantengo fechas como string porque el formato es "dd-MM-yyyy" en tu ejemplo.
        [JsonProperty("fecha_de_emision")]
        public string FechaDeEmision { get; set; }

        [JsonProperty("observaciones")]
        public string Observaciones { get; set; }

        [JsonProperty("motivo_de_traslado")]
        public string MotivoDeTraslado { get; set; }

        [JsonProperty("peso_bruto_total")]
        public string PesoBrutoTotal { get; set; }

        [JsonProperty("peso_bruto_unidad_de_medida")]
        public string PesoBrutoUnidadDeMedida { get; set; }

        [JsonProperty("numero_de_bultos")]
        public string NumeroDeBultos { get; set; }

        [JsonProperty("tipo_de_transporte")]
        public string TipoDeTransporte { get; set; }

        [JsonProperty("fecha_de_inicio_de_traslado")]
        public string FechaDeInicioDeTraslado { get; set; }

        [JsonProperty("transportista_documento_tipo")]
        public string TransportistaDocumentoTipo { get; set; }

        [JsonProperty("transportista_documento_numero")]
        public string TransportistaDocumentoNumero { get; set; }

        [JsonProperty("transportista_denominacion")]
        public string TransportistaDenominacion { get; set; }

        [JsonProperty("transportista_placa_numero")]
        public string TransportistaPlacaNumero { get; set; }

        [JsonProperty("conductor_documento_tipo")]
        public string ConductorDocumentoTipo { get; set; }

        [JsonProperty("conductor_documento_numero")]
        public string ConductorDocumentoNumero { get; set; }

        [JsonProperty("conductor_nombre")]
        public string ConductorNombre { get; set; }

        [JsonProperty("conductor_apellidos")]
        public string ConductorApellidos { get; set; }

        [JsonProperty("conductor_numero_licencia")]
        public string ConductorNumeroLicencia { get; set; }

        [JsonProperty("punto_de_partida_ubigeo")]
        public string PuntoDePartidaUbigeo { get; set; }

        [JsonProperty("punto_de_partida_direccion")]
        public string PuntoDePartidaDireccion { get; set; }

        [JsonProperty("punto_de_partida_codigo_establecimiento_sunat")]
        public string PuntoDePartidaCodigoEstablecimientoSunat { get; set; }

        [JsonProperty("punto_de_llegada_ubigeo")]
        public string PuntoDeLlegadaUbigeo { get; set; }

        [JsonProperty("punto_de_llegada_direccion")]
        public string PuntoDeLlegadaDireccion { get; set; }

        [JsonProperty("punto_de_llegada_codigo_establecimiento_sunat")]
        public string PuntoDeLlegadaCodigoEstablecimientoSunat { get; set; }

        [JsonProperty("enviar_automaticamente_al_cliente")]
        public string EnviarAutomaticamenteAlCliente { get; set; }

        [JsonProperty("formato_de_pdf")]
        public string FormatoDePdf { get; set; }

        [JsonProperty("items")]
        public List<GuiaItem> Items { get; set; }

        [JsonProperty("documento_relacionado")]
        public List<DocumentoRelacionado> DocumentoRelacionado { get; set; }

        [JsonProperty("vehiculos_secundarios")]
        public List<VehiculoSecundario> VehiculosSecundarios { get; set; }

        [JsonProperty("conductores_secundarios")]
        public List<ConductorSecundario> ConductoresSecundarios { get; set; }
    }

    public class GuiaItem
    {
        [JsonProperty("unidad_de_medida")]
        public string UnidadDeMedida { get; set; }

        [JsonProperty("codigo")]
        public string Codigo { get; set; }

        [JsonProperty("descripcion")]
        public string Descripcion { get; set; }

        [JsonProperty("cantidad")]
        public string Cantidad { get; set; }
    }

    public class DocumentoRelacionado
    {
        [JsonProperty("tipo")]
        public string Tipo { get; set; }

        [JsonProperty("serie")]
        public string Serie { get; set; }

        [JsonProperty("numero")]
        public string Numero { get; set; }
    }

    public class VehiculoSecundario
    {
        [JsonProperty("placa_numero")]
        public string PlacaNumero { get; set; }
    }

    public class ConductorSecundario
    {
        [JsonProperty("documento_tipo")]
        public string DocumentoTipo { get; set; }

        [JsonProperty("documento_numero")]
        public string DocumentoNumero { get; set; }

        [JsonProperty("nombre")]
        public string Nombre { get; set; }

        [JsonProperty("apellidos")]
        public string Apellidos { get; set; }

        [JsonProperty("numero_licencia")]
        public string NumeroLicencia { get; set; }

    }
}
