using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;
using System.Threading.Tasks;
using RightNow.AddIns.AddInViews;
using System.Web.Script.Serialization;
using static Parts_Required.WebServiceReqdParams;

namespace Parts_Required
{
    class PartsRequiredModel
    {
        string _xmlnsURL;
        string _curlURL;
        string _headerURL;
        string _responsibility;
        string _respApplication;
        string _securityGroup;
        string _orgID;
        string _nlsLanguage;

        POEHEADERREC _partsHeaderRecord = new POEHEADERREC();
        List<OELINEREC> _lineRecords;
        IIncident _incidentRecord;
        public static RightNowConnectService _rnConnectService;

        public PartsRequiredModel(IRecordContext RecordContext)
        {
            _rnConnectService = RightNowConnectService.GetService();

            string partsRequiredConfigValue = _rnConnectService.GetConfigValue("CUSTOM_CFG_PARTS_REQUIRED");
            if (partsRequiredConfigValue != null)
            {
                var s = new JavaScriptSerializer();
                var configVerb = s.Deserialize<WebServiceConfigVerbs.RootObject>(partsRequiredConfigValue);
                _curlURL = configVerb.URL;
                _xmlnsURL = configVerb.xmlns;
                _headerURL = configVerb.RESTHeader.xmlns;
                _respApplication = configVerb.RESTHeader.RespApplication;
                _responsibility = configVerb.RESTHeader.Responsibility;
                _securityGroup = configVerb.RESTHeader.SecurityGroup;
                _nlsLanguage = configVerb.RESTHeader.NLSLanguage;
                _orgID = configVerb.RESTHeader.Org_Id;
            }
        }

        /// <summary>
        /// Get required details to build WebRequest
        /// </summary>
        /// <param name="partNumber"></param>
        /// <param name="partQty"></param>
        /// <param name="incidentRecord"></param>
        /// <param name="partsOrderID"></param>
        public void GetDetails(List<OELINEREC> lineRecords, IIncident incidentRecord, int numOfVIN, POEHEADERREC partsHeaderRecord,
                               int shipToSiteID, int partOdrInstrID)
        {
            _incidentRecord = incidentRecord;
            _lineRecords = lineRecords;
            _partsHeaderRecord = partsHeaderRecord;

            //If all required info is valid then form jSon request parameter
            var content = GetReqParam();
            var jsonContent = WebServiceRequest.JsonSerialize(content);
            jsonContent = jsonContent.Replace("xmlns", "@xmlns");

            //Call webservice                 
            string jsonResponse = WebServiceRequest.Get(_curlURL, jsonContent, "POST");
            if (jsonResponse == "")
            {
                //Destroy the partsorder objects
                _rnConnectService.DestroyPartsOrder(_lineRecords);
                ReportCommandAddIn.form.Hide();
                MessageBox.Show("Server didn't return any info");
            }
            else
            {
                ExtractResponse(jsonResponse, numOfVIN, shipToSiteID, partOdrInstrID);
            }            

        }

        /// <summary>
        /// Function to get Request param need for web-service
        /// </summary>
        public RootObject GetReqParam( )
        {
            RootObject content = new RootObject
            {
                CREATE_A_SALES_ORDER_Input = new CREATEASALESORDERInput
                {
                    xmlns = _xmlnsURL,
                    RESTHeader = new RESTHeader
                    {
                        xmlns = _headerURL,
                        Responsibility = _responsibility,
                        RespApplication = _respApplication,
                        SecurityGroup = _securityGroup,
                        NLSLanguage = _nlsLanguage,
                        Org_Id = _orgID
                    },
                    InputParameters = new InputParameters
                    {
                        P_OE_HEADER_REC = _partsHeaderRecord,
                        P_OE_LINE_TBL = new POELINETBL
                        {
                            OE_LINE_REC = _lineRecords
                        }
                    }
                }
            };

            return content;
        }

        /// <summary>
        /// Funtion to handle ebs webservice response
        /// </summary>
        /// <param name="respJson">response in jSON string</param>
        public void ExtractResponse(string jsonResp, int numOfVIN, int shipToSiteID, int partOdrInstrID)
        {
            Dictionary<String, object> retunTblItem;
            //Extract response
            var deserializeJson = new JavaScriptSerializer().Deserialize<dynamic>(jsonResp);
            Dictionary<String, object> data = (Dictionary<string, object>)deserializeJson;
            Dictionary<String, object> outputParameters = (Dictionary<String, object>)data["OutputParameters"];
            Dictionary<String, object> returnTbl = (Dictionary<String, object>)outputParameters["P_RETURN_TBL"];
            //Check if its an array, is so extract first element
            if (IsArray(returnTbl["P_RETURN_TBL_ITEM"]))
            {
                object[] tblItems = (object[])returnTbl["P_RETURN_TBL_ITEM"];
                retunTblItem = (Dictionary<string, object>)tblItems[0];
            }
            else
            {
                retunTblItem = (Dictionary<String, object>)returnTbl["P_RETURN_TBL_ITEM"];
            }

            if (retunTblItem["HASERROR"].ToString() == "0")//If Success
            {
                //if (partsOrderID != 0)
                //{
                    string salesOrderNum = retunTblItem["SALES_ORDER_NO"].ToString();

                    //Update PartsOrder Object with SalesOrder Number and other Reported Incident Info
                    _rnConnectService.UpdatePartsOrder(_lineRecords, salesOrderNum, numOfVIN, _partsHeaderRecord, shipToSiteID, partOdrInstrID);
                    
                    //Clear out some of Reported Incident field that were specific for this order only for Retrofit case only
                    if(partOdrInstrID >0)
                    {
                        IRecordContext recordContext = ReportCommandAddIn._globalContext.AutomationContext.CurrentWorkspace;
                        recordContext.TriggerNamedEvent("clearPartsOrderFields");
                    }
                //}
            } 
            else
            {
                //Destroy the partsorder objects
                _rnConnectService.DestroyPartsOrder(_lineRecords);
                ReportCommandAddIn.form.Hide();
                MessageBox.Show(retunTblItem["DESCRIPTION"].ToString());
            }           
        }
        /// <summary>
        /// Funtion to check if object is array
        /// </summary>
        /// <param name="bool">response in boolean</param>
        public static bool IsArray(object o)
        {
            return o is Array;
        }
    }
}
