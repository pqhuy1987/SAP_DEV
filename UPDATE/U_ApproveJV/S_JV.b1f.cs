using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Xml;
using System.Data;

namespace U_ApproveJV
{
    [FormAttribute("229", "S_JV.b1f")]
    class S_JV : SystemFormBase
    {
        public S_JV()
        {
        }
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("8").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("4").Specific));
            this.Matrix0.PressedAfter += new SAPbouiCOM._IMatrixEvents_PressedAfterEventHandler(this.Matrix0_PressedAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.DataAddAfter += new SAPbouiCOM.Framework.FormBase.DataAddAfterHandler(this.Form_DataAddAfter);
            this.DataDeleteBefore += new SAPbouiCOM.Framework.FormBase.DataDeleteBeforeHandler(this.Form_DataDeleteBefore);
            this.DataUpdateBefore += new DataUpdateBeforeHandler(this.Form_DataUpdateBefore);

        }

        private void Form_DataAddAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            //throw new System.NotImplementedException();
            if (pVal.ActionSuccess)
            {
                //Get Batch Num has created
                XmlDocument xmldoc = new XmlDocument();
                xmldoc.LoadXml(((SAPbouiCOM.BusinessObjectInfo)pVal).ObjectKey);
                XmlNodeList nodeList = xmldoc.GetElementsByTagName("BatchNum");
                string BatchNum = string.Empty;
                try
                {
                    if (nodeList.Count > 0)
                        BatchNum = nodeList.Item(0).InnerText;
                }
                catch
                {
                    BatchNum = string.Empty;
                }
                if (!string.IsNullOrEmpty(BatchNum))
                {

                    //Check if exist in JV_APPROVE
                    if (!Check_Approve_Process_Exist(BatchNum))
                    {
                        //Get Info BatchNum
                        SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oR_RecordSet.DoQuery("Select top 1 U_LCP,Project from OBTF where BatchNum=" + BatchNum + " order by TransID desc");
                        string U_LCP = oR_RecordSet.Fields.Item("U_LCP").Value.ToString();
                        //string BCH = oR_RecordSet.Fields.Item("U_CP").Value.ToString();
                        string FProject = oR_RecordSet.Fields.Item("Project").Value.ToString();
                        string Type_JV = "";
                        DataTable tb_lstappr = null;
                        if (U_LCP == "VP")
                        {
                            Type_JV = "VP";
                        }
                        else if (U_LCP == "BCH")
                        {
                            Type_JV = "BCH";
                        }
                        if (Type_JV == "VP" || Type_JV == "BCH")
                        {
                            if (Check_Accountant())

                                Button0.Item.Visible = false;
                        }
                        else
                        {
                            if (Check_Accountant())
                                Button0.Item.Visible = true;
                        }
                        tb_lstappr = GetList_Approve(FProject, Type_JV);
                        if (tb_lstappr.Rows.Count > 0)
                        {
                            SAPbobsCOM.GeneralService oGeneralService = null;
                            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                            SAPbobsCOM.CompanyService sCmp = null;
                            SAPbobsCOM.GeneralData oGeneralData = null;
                            SAPbobsCOM.GeneralData oChild = null;
                            SAPbobsCOM.GeneralDataCollection oChildren = null;
                            sCmp = oCompany.GetCompanyService();
                            oGeneralService = sCmp.GetGeneralService("JVAPPROVE");
                            //Create data for new row in main UDO
                            oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                            oGeneralData.SetProperty("U_JVBatchNum", int.Parse(BatchNum));
                            oGeneralData.SetProperty("U_Type", Type_JV);
                            //Create data for a row in the child table
                            oChildren = oGeneralData.Child("JV_APROVE_D");
                            foreach (DataRow r in tb_lstappr.Rows)
                            {
                                oChild = oChildren.Add();
                                oChild.SetProperty("U_Level", r["LEVEL"].ToString());
                                oChild.SetProperty("U_Position", r["Position"].ToString());
                            }
                            oGeneralParams = oGeneralService.Add(oGeneralData);
                            if (!String.IsNullOrEmpty(oGeneralParams.GetProperty("DocEntry").ToString()))
                            {
                                oApp.SetStatusBarMessage("Approve Process Added !!!", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                if (Check_Accountant())
                                    Button0.Item.Visible = false;
                            }
                            else
                            {
                                //Close Voucher
                                //SAPbobsCOM.JournalVouchers oVoucher = (SAPbobsCOM.JournalVouchers)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers);
                            }
                        }
                    }
                }
            }
        }

        private void OnCustomInitialize()
        {
            this.oApp = (SAPbouiCOM.Application)Application.SBO_Application;
            this.oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            //Create Connection SQL
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery("Select * from [@ADDONCFG]");
            if (oR_RecordSet.RecordCount > 0)
            {
                string uid = oR_RecordSet.Fields.Item("Code").Value.ToString();
                string pwd = oR_RecordSet.Fields.Item("Name").Value.ToString();
                conn = new SqlConnection(string.Format("Data Source={0}; Initial Catalog={1}; User id={2}; Password={3};", oCompany.Server, oCompany.CompanyDB, uid, pwd));
            }
            else
            {
                oApp.MessageBox("Can't connect DB !");
            }
        }

        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Button Button0;

        private void Matrix0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            int BatchNum = 0;
            int.TryParse(((SAPbouiCOM.EditText)(Matrix0.Columns.Item("1").Cells.Item(pVal.Row).Specific)).Value,out BatchNum);
            DataTable tb_lst = Check_BCH_VP(BatchNum);
            string Lastest_Approve ="";
            string Type ="";
            if (tb_lst.Rows.Count > 0)
            {
                foreach (DataRow r in tb_lst.Rows)
                {
                    Lastest_Approve = r["APPROVED"].ToString();
                    Type = r["U_LCP"].ToString();
                }
            }
            if (Type == "BCH" || Type == "VP")
            {
                if (Check_Accountant())
                {
                    if (Lastest_Approve == "1")
                        Button0.Item.Visible = true;
                    else
                        Button0.Item.Visible = false;
                }
            }
            else
            {
                if (Check_Accountant())
                {
                    Button0.Item.Visible = true;
                }
            }
        }

        private bool Check_Accountant()
        {
            string sql = string.Format("Select dept from OHEM  where userID = (Select t.USERID from OUSR t where t.User_Code='{0}')", oCompany.UserName);
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(sql);
            if (oR_RecordSet.RecordCount > 0)
            {
                string dept = oR_RecordSet.Fields.Item("dept").Value.ToString();
                if (dept == "-2") return true;
                else return false;
            }
            return false;
        }

        private bool Check_Approve_Process_Exist(string pBatchNum)
        {
            string sql = string.Format("Select Count(*) as A from [@JV_APPROVE] where U_JVBatchNum = {0}", pBatchNum);
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(sql);
            if (oR_RecordSet.RecordCount > 0)
            {
                string tmp = oR_RecordSet.Fields.Item("A").Value.ToString();
                if (tmp == "0") return false;
                else return true;
            }
            return false;
        }

        System.Data.DataTable GetList_Approve(string pFProject, string pType)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("JV_GetList_Approve", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialPrj", pFProject);
                cmd.Parameters.AddWithValue("@Usr", oCompany.UserName);
                cmd.Parameters.AddWithValue("@Type", pType);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
            return result;
        }

        System.Data.DataTable Check_BCH_VP(int pBatchNum)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("JV_Check_BCH_VP", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@BatchNum", pBatchNum);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
            return result;
        }

        private void Form_DataDeleteBefore(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            oApp.MessageBox("Delete Event");

        }

        private void Form_DataUpdateBefore(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            oApp.MessageBox("Update Event");

        }

    }
}
