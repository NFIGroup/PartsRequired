using System;
using System.AddIn;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using RightNow.AddIns.Common;
using static Parts_Required.WebServiceReqdParams;

namespace Parts_Required
{
    [AddIn("Report Command AddIn", Version = "1.0.0.0")]
    public class ReportCommandAddIn : IReportCommand2
    {
        static public IGlobalContext _globalContext;
        IRecordContext _recordContext;
        IIncident _incidentRecord;
        IList<IReportRow> _selectedRows;
        private IGenericObject _incidentExtra;
        public static ProgressForm form = new ProgressForm();

        #region IReportCommand Members
        public string Text
        {
            get
            {
                return "Order Parts";
            }
        }

        public string Tooltip
        {
            get
            {
                return "Call Parts Required Webservice to send parts order";
            }
        }
        public Image Image16
        {
            get
            {
                return Properties.Resources.AddIn16;
            }
        }
        public Image Image32
        {
            get
            {
                return Properties.Resources.AddIn32;
            }
        }

        public IList<ReportRecordIdType> RecordTypes
        {
            get
            {
                IList<ReportRecordIdType> typeList = new List<ReportRecordIdType>();

                typeList.Add(ReportRecordIdType.Answer);
                typeList.Add(ReportRecordIdType.Chat);
                typeList.Add(ReportRecordIdType.CloudAcct2Search);
                typeList.Add(ReportRecordIdType.Contact);
                typeList.Add(ReportRecordIdType.ContactList);
                typeList.Add(ReportRecordIdType.Document);
                typeList.Add(ReportRecordIdType.Flow);
                typeList.Add(ReportRecordIdType.Incident);
                typeList.Add(ReportRecordIdType.Mailing);
                typeList.Add(ReportRecordIdType.MetaAnswer);
                typeList.Add(ReportRecordIdType.Opportunity);
                typeList.Add(ReportRecordIdType.Organization);
                typeList.Add(ReportRecordIdType.Question);
                typeList.Add(ReportRecordIdType.QueuedReport);
                typeList.Add(ReportRecordIdType.Quote);
                typeList.Add(ReportRecordIdType.QuoteProduct);
                typeList.Add(ReportRecordIdType.Report);
                typeList.Add(ReportRecordIdType.Segment);
                typeList.Add(ReportRecordIdType.Survey);
                typeList.Add(ReportRecordIdType.Task);
                typeList.Add(ReportRecordIdType.CustomObjectAll);
                return typeList;
            }
        }
        public IList<string> CustomObjectRecordTypes
        {
            get
            {
                IList<string> typeList = new List<string>();

                typeList.Add("PartsOrder");
                typeList.Add("Parts_PartsOrder");
                typeList.Add("Parts");
                return typeList;
            }
        }
        public bool Enabled(IList<IReportRow> rows)
        {
            return true;
        }

        public void Execute(IList<IReportRow> rows)
        {
            _selectedRows = rows;
            _recordContext = _globalContext.AutomationContext.CurrentWorkspace;
            _incidentRecord = _recordContext.GetWorkspaceRecord(WorkspaceRecordType.Incident) as IIncident;
            _incidentExtra = _recordContext.GetWorkspaceRecord("CO$Incident_Extra") as IGenericObject;

            System.Threading.Thread th = new System.Threading.Thread(ProcessSelectedRowInfo);
            th.Start();
            form.Show();
        }

        /// <summary>
        /// To get Parts details and Call Webservice
        /// </summary>
        public void ProcessSelectedRowInfo()
        {
            POEHEADERREC partsHeaderRecord = new POEHEADERREC();
            RightNowConnectService _rnConnectService = RightNowConnectService.GetService();
            /*Get Reported Incident Info*/

            //Get Incident Type
            string incType = _rnConnectService.getIncidentField("c", "Incident Type", _incidentRecord);

            //Get ship set from Current Reported Incident
            string shipSet = _rnConnectService.getIncidentField("CO", "ship_set", _incidentRecord);

            //Get Num of VIN for which this order will be placed
            string numOfVINInString = _rnConnectService.getIncidentField("CO", "no_of_vins", _incidentRecord);
            if (numOfVINInString == String.Empty)
            {
                if (incType == "30")//case of claim type
                {
                    numOfVINInString = "1";//set it to 1
                }
                else
                {
                    form.Hide();
                    MessageBox.Show("Number Of VIN is empty");
                    return;
                }
            }
            int numOfVIN = Convert.ToInt32(numOfVINInString);

            string orderTypeId = _rnConnectService.getIncidentField("CO", "order_type", _incidentRecord);
            if (orderTypeId == String.Empty)
            {
                form.Hide();
                MessageBox.Show("Order Type is empty");
                return;
            }
            partsHeaderRecord.ORDER_TYPE = _rnConnectService.GetOrderTypeName(Convert.ToInt32(orderTypeId));

            string shipToId = _rnConnectService.getIncidentField("CO", "Ship_to_site", _incidentRecord);
            if (shipToId == String.Empty)
            {
                form.Hide();
                MessageBox.Show("Ship to Site is empty");
                return;
            }
            partsHeaderRecord.SHIP_TO_ORG_ID = Convert.ToInt32(_rnConnectService.GetEbsID(Convert.ToInt32(shipToId)));

            string shipDate = _rnConnectService.getIncidentField("CO", "requested_ship_date", _incidentRecord);
            if (shipDate == String.Empty)
            {
                if (incType == "30")//case of claim
                {
                    shipDate = DateTime.Today.ToString();//set to today's date
                }
                else
                {
                    form.Hide();
                    MessageBox.Show("Requested Ship Date is empty");
                    return;
                }
            }
            partsHeaderRecord.SHIP_DATE = Convert.ToDateTime(shipDate).ToString("dd-MMM-yyyy");

            string dcID = _rnConnectService.GetIncidentExtraFieldValue("Distribution_center", _incidentExtra);
            if (dcID == String.Empty)
            {
                form.Hide();
                MessageBox.Show("Distribution Center is empty");
                return;
            }
            partsHeaderRecord.WAREHOUSE_ORG = _rnConnectService.GetDistributionCenterIDName(Convert.ToInt32(dcID));

            partsHeaderRecord.CUSTOMER_ID = Convert.ToInt32(_rnConnectService.GetCustomerEbsID((int)_incidentRecord.OrgID));
            partsHeaderRecord.CLAIM_NUMBER = _incidentRecord.RefNo;
            partsHeaderRecord.PROJECT_NUMBER = _rnConnectService.getIncidentField("CO", "project_number", _incidentRecord);
            partsHeaderRecord.RETROFIT_NUMBER = _rnConnectService.getIncidentField("CO", "retrofit_number", _incidentRecord);
            partsHeaderRecord.CUST_PO_NUMBER = _rnConnectService.getIncidentField("CO", "PO_Number", _incidentRecord);
            partsHeaderRecord.SHIPPING_INSTRUCTIONS = _rnConnectService.getIncidentField("CO", "shipping_instructions", _incidentRecord);

            string partOdrInstrIDString = _rnConnectService.getIncidentField("CO", "PartOrderInstruction_ID", _incidentRecord);
            int partOdrInstrID = 0;//default case where we won't have partsorderinstruction
            if (partOdrInstrIDString != String.Empty)
            {
                partOdrInstrID = Convert.ToInt32(partOdrInstrIDString);
            }

            //List<int> partsIDOrdered = new List<int>();
            List<OELINEREC> lineRecords = new List<OELINEREC>();

            //Loop over each selected parts that user wants to order
            //Get required info needed for EBS web-service
            double oq = 0.00; 
            foreach (IReportRow row in _selectedRows)
            {
                OELINEREC lineRecord = new OELINEREC();
                IList<IReportCell> cells = row.Cells;

                foreach (IReportCell cell in cells)
                {
                    if (cell.Name == "Part ID")
                    {
                        //Create PartsOrder record first as EBS web-service need PartsOrder ID
                        lineRecord.ORDERED_ID = _rnConnectService.CreatePartsOrder(Convert.ToInt32(cell.Value));//Pass parts Order ID
                    }
                    if (cell.Name == "Part #")
                    {
                        lineRecord.ORDERED_ITEM = cell.Value;
                    }
                    if (cell.Name == "Qty")
                    {
                        oq = Convert.ToDouble(Math.Round(Convert.ToDecimal(cell.Value), 2).ToString()) * numOfVIN;
                        lineRecord.ORDERED_QUANTITY = Math.Round(oq, 2);//get total qyantity 
                    }
                    if (cell.Name == "Source Type")
                    {
                        lineRecord.SOURCE_TYPE = cell.Value;
                    }
                }
                lineRecord.SHIP_SET = shipSet;

                lineRecords.Add(lineRecord);
            }
            //Call PartsRequired Model to send parts info to EBS
            PartsRequiredModel partsRequired = new PartsRequiredModel(_recordContext);
            partsRequired.GetDetails(lineRecords, _incidentRecord, numOfVIN, partsHeaderRecord, Convert.ToInt32(shipToId), partOdrInstrID);
            form.Hide();
            _recordContext.ExecuteEditorCommand(RightNow.AddIns.Common.EditorCommand.Save);
        }
        #endregion

        #region IAddInBase Members

        /// <summary>
        /// Method which is invoked from the Add-In framework and is used to programmatically control whether to load the Add-In.
        /// </summary>
        /// <param name="GlobalContext">The Global Context for the Add-In framework.</param>
        /// <returns>If true the Add-In to be loaded, if false the Add-In will not be loaded.</returns>
        public bool Initialize(IGlobalContext context)
        {
            _globalContext = context;
            return true;
        }
        #endregion
    }
}
