using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Parts_Required
{
    class WebServiceConfigVerbs
    {
        public class RESTHeader
        {
            public string xmlns { get; set; }
            public string Responsibility { get; set; }
            public string RespApplication { get; set; }
            public string SecurityGroup { get; set; }
            public string NLSLanguage { get; set; }
            public string Org_Id { get; set; }
        }

        public class RootObject
        {
            public string URL { get; set; }
            public string xmlns { get; set; }
            public RESTHeader RESTHeader { get; set; }
        }
    }
}
