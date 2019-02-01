using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Parts_Required
{
    class WebServiceReqdParams
    {
        public class RESTHeader
        {
            public string @xmlns { get; set; }
            public string Responsibility { get; set; }
            public string RespApplication { get; set; }
            public string SecurityGroup { get; set; }
            public string NLSLanguage { get; set; }
            public string Org_Id { get; set; }
        }

        public class POEHEADERREC
        {
            public string ORDER_TYPE { get; set; }
            public int CUSTOMER_ID { get; set; }
            public int SHIP_TO_ORG_ID { get; set; }
            public string CLAIM_NUMBER { get; set; }
            public string PROJECT_NUMBER { get; set; }
            public string RETROFIT_NUMBER { get; set; }
            public string CUST_PO_NUMBER { get; set; }
            public string WAREHOUSE_ORG { get; set; }
            public string SHIP_DATE { get; set; }
            public string SHIPPING_INSTRUCTIONS { get; set; }
        }

        public class OELINEREC
        {
            public string ORDERED_ITEM { get; set; }
            public int ORDERED_ID { get; set; }
            public double ORDERED_QUANTITY { get; set; }
            public string SOURCE_TYPE { get; set; }
            public string SHIP_SET { get; set; }
        }

        public class POELINETBL
        {
            public List<OELINEREC> OE_LINE_REC { get; set; }
        }

        public class InputParameters
        {
            public POEHEADERREC P_OE_HEADER_REC { get; set; }
            public POELINETBL P_OE_LINE_TBL { get; set; }
        }

        public class CREATEASALESORDERInput
        {
            public string @xmlns { get; set; }
            public RESTHeader RESTHeader { get; set; }
            public InputParameters InputParameters { get; set; }
        }

        public class RootObject
        {
            public CREATEASALESORDERInput CREATE_A_SALES_ORDER_Input { get; set; }
        }
    }
}
