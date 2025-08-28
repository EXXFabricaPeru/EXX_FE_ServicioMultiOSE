using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServicioConectorFE.Entities.SAP
{
    public class CONF
    {
        public string IdEmpresa { get; set; }
        public int DbServerType { get; set; }
        public string Server { get; set; }
        public string CompanyDB { get; set; }
        public string DbUserName { get; set; }
        public string DbPassword { get; set; }
        public string LicenseServer { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
    }
}
