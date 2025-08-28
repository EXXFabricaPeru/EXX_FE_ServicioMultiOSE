using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ServicioConectorFE.Entities.SAP
{
    public class OADM
    {
        public String MainCurncy { get; set; }
        public String SysCurrncy { get; set; }
        public String CompnyName { get; set; }
        public String PrintHeader { get; set; }
        public String CompnyAddr { get; set; }
        public String Country { get; set; }
        public String TaxIdNum { get; set; }
        public String RevOffice { get; set; }
        public String AliasName { get; set; }
        public String E_Mail { get; set; }
        public String Phone1 { get; set; }
        public String Phone2 { get; set; }
        public String Fax { get; set; }
    }

    public class ADM1
    {
        public String ZipCode { get; set; }
        public String Building { get; set; }
        public String County { get; set; }
        public String City { get; set; }
        public String State { get; set; }
        public String Street { get; set; }
        public String IntrntAdrs { get; set; }
    }
}
