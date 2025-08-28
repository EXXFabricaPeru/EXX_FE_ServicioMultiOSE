using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServicioConectorFE.Entities.SAP
{
    public class Baja
    {
        public string Tabla { get; set; }
        public int ObjType { get; set; }
        public int DocEntry { get; set; }
        public int ObjTypeAux { get; set; }
        public int DocEntryAux { get; set; }
        public string Estado { get; set; }
        public string tipoComprobante { get; set; }
        public string serie { get; set; }
        public string correlativo { get; set; }
        public string fechaEmision { get; set; }
        public string notaSustento { get; set; }
    }
}
