using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServicioConectorFE.Entities.SAP
{
    public class Comprobante
    {
        public string Tabla { get; set; }
        public int ObjType { get; set; }
        public int DocEntry { get; set; }
        public string NumeroSAP { get; set; }

        public string Ruc { get; set; }
        public string RazonSocial { get; set; }
        public string TipoDocumento { get; set; }
        public string Serie { get; set; }
        public string Correlativo { get; set; }
        public string CodTipoDoc { get; set; }
        public string Moneda { get; set; }
        public string Fecha { get; set; }
        public string Estado { get; set; }
        public string Impuesto { get; set; }
        public string ImporteTotal { get; set; }

        public string Nota { get; set; }
        public string EntryCodTipo { get; set; }
        public int ObjectId { get; internal set; }
        public string ReferenciaSerie { get; set; }
        public string Sucursal { get; set; }
    }
}
