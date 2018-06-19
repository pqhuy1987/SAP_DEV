using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;

namespace Approve_PR
{
    [FormAttribute("Approve_PR.PR_Aprrove", "PR_Aprrove.b1f")]
    class PR_Aprrove : UserFormBase
    {
        public PR_Aprrove()
        {
        }
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        int DocEntry_ODRF = 0;
        int WddCode = 0;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        
        private SAPbouiCOM.OptionBtn OptionBtn0;
        private SAPbouiCOM.OptionBtn OptionBtn1;
        private SAPbouiCOM.OptionBtn OptionBtn2;
        private SAPbouiCOM.OptionBtn OptionBtn3;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Grid Grid1;
        private SAPbouiCOM.Grid Grid2;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;

        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("grd_lst").Specific));
            this.Grid0.PressedAfter += new SAPbouiCOM._IGridEvents_PressedAfterEventHandler(this.Grid0_PressedAfter);
            this.OptionBtn0 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_cur").Specific));
            this.OptionBtn0.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn0_PressedAfter);
            this.OptionBtn1 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_appr").Specific));
            this.OptionBtn1.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn1_PressedAfter);
            this.OptionBtn2 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_rej").Specific));
            this.OptionBtn2.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn2_PressedAfter);
            this.OptionBtn3 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_all").Specific));
            this.OptionBtn3.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn3_PressedAfter);
            this.Grid1 = ((SAPbouiCOM.Grid)(this.GetItem("grd_pro").Specific));
            this.Grid2 = ((SAPbouiCOM.Grid)(this.GetItem("grd_info").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_comm").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_appr").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_rej").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>

        public override void OnInitializeFormEvents()
        {
            
        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {
            this.oApp = (SAPbouiCOM.Application)SAPbouiCOM.Framework.Application.SBO_Application;
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

            OptionBtn1.GroupWith("op_cur");
            OptionBtn2.GroupWith("op_cur");
            OptionBtn3.GroupWith("op_cur");
            OptionBtn0.Selected = true;
        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            int new_heigt = this.UIAPIRawForm.ClientHeight;
            int new_width = this.UIAPIRawForm.ClientWidth;

            //Resize List JV
            Grid0.Item.Width = 400;
            Grid0.Item.Height = new_heigt - 40;
            //Resize List Apprv
            Grid1.Item.Left = 422;
            Grid1.Item.Height = 130;

            //Resize List Infor
            Grid2.Item.Left = 422;

            //Resize Button
            Button0.Item.Top = new_heigt - 40;
            Button0.Item.Left = 422;
            Button1.Item.Top = new_heigt - 40;
            Button1.Item.Left = 588;

            //Resize option button
            OptionBtn1.Item.Left = OptionBtn0.Item.Left + OptionBtn0.Item.Width + 10;
            OptionBtn2.Item.Left = OptionBtn1.Item.Left + OptionBtn1.Item.Width + 10;
            OptionBtn3.Item.Left = OptionBtn2.Item.Left + OptionBtn2.Item.Width + 10;

        }

        private bool CheckExistUniqueID(SAPbouiCOM.Form pForm, string pItemID)
        {
            if (pForm.DataSources.DataTables.Count > 0)
            {
                for (int i = 0; i < pForm.DataSources.DataTables.Count; i++)
                {
                    if (pForm.DataSources.DataTables.Item(i).UniqueID == pItemID)
                    {
                        return true;
                    }
                }
                return false;
            }
            else
            {
                return false;
            }
        }

        private SAPbouiCOM.DataTable Convert_SAP_DataTable(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_PRList"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_PRList");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_PRList");
            }

            //Add column to DataTable
            foreach (System.Data.DataColumn c in pDataTable.Columns)
            {
                try
                {
                    if (c.DataType.ToString() == "System.DateTime")
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Date);
                    else
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Text);
                }
                catch
                { }

            }
            //Add row to DataTable
            foreach (System.Data.DataRow r in pDataTable.Rows)
            {
                oDT.Rows.Add();
                foreach (System.Data.DataColumn c in pDataTable.Columns)
                {
                    oDT.SetValue(c.ColumnName, oDT.Rows.Count - 1, r[c.ColumnName].ToString());
                }
            }

            return oDT;
        }

        private SAPbouiCOM.DataTable Convert_SAP_DataTable_Process(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_PRProcess"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_PRProcess");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_PRProcess");
            }

            //Add column to DataTable
            foreach (System.Data.DataColumn c in pDataTable.Columns)
            {
                try
                {
                    if (c.DataType.ToString() == "System.DateTime")
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Date);
                    else
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Text);
                }
                catch
                { }

            }
            //Add row to DataTable
            foreach (System.Data.DataRow r in pDataTable.Rows)
            {
                oDT.Rows.Add();
                foreach (System.Data.DataColumn c in pDataTable.Columns)
                {
                    oDT.SetValue(c.ColumnName, oDT.Rows.Count - 1, r[c.ColumnName].ToString());
                }
            }

            return oDT;
        }

        private SAPbouiCOM.DataTable Convert_SAP_DataTable_Details(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_PRDetails"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_PRDetails");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_PRDetails");
            }

            //Add column to DataTable
            foreach (System.Data.DataColumn c in pDataTable.Columns)
            {
                try
                {
                    if (c.DataType.ToString() == "System.DateTime")
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Date);
                    else
                        oDT.Columns.Add(c.ColumnName, SAPbouiCOM.BoFieldsType.ft_Text);
                }
                catch
                { }

            }
            //Add row to DataTable
            foreach (System.Data.DataRow r in pDataTable.Rows)
            {
                oDT.Rows.Add();
                foreach (System.Data.DataColumn c in pDataTable.Columns)
                {
                    oDT.SetValue(c.ColumnName, oDT.Rows.Count - 1, r[c.ColumnName].ToString());
                }
            }

            return oDT;
        }
        //Grid Approve
        private void Load_Grid_Period()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("PR_Get_List_Approve", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
                conn.Close();
                cmd.Dispose();
            }
        }

        //Grid Approved
        private void Load_Grid_Approved_Period()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("PR_Get_List_Approved", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
                conn.Close();
                cmd.Dispose();
            }
        }

        //Grid Rejected
        private void Load_Grid_Rejected_Period()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("PR_Get_List_Rejected", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
                conn.Close();
                cmd.Dispose();
            }
        }

        //Grid All
        private void Load_Grid_All_Period()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("[PR_Get_List_All]", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
                conn.Close();
                cmd.Dispose();
            }
        }

        //Grid Approve Process
        private void Load_Grid_Process()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("PR_Get_Approve_Process", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry_ODRF", DocEntry_ODRF);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid1.DataTable = Convert_SAP_DataTable_Process(result);
                Grid1.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
                conn.Close();
                cmd.Dispose();
            }
        }

        //Grid Document Details
        private void Load_Grid_Details()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("PR_Get_Document_Details", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry_ODRF", DocEntry_ODRF);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid2.DataTable = Convert_SAP_DataTable_Details(result);
                Grid2.AutoResizeColumns();

            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
                conn.Close();
                cmd.Dispose();
            }
        }

        //Show Current
        private void OptionBtn0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn0.Selected)
            {
                Load_Grid_Period();
                Button0.Item.Enabled = true;
                Button1.Item.Enabled = true;
            }
        }

        //Show Approved
        private void OptionBtn1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn1.Selected)
            {
                Load_Grid_Approved_Period();
                Button0.Item.Enabled = false;
                Button1.Item.Enabled = false;
            }
        }

        //Show Rejected
        private void OptionBtn2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn2.Selected)
            {
                Load_Grid_Rejected_Period();
                Button0.Item.Enabled = false;
                Button1.Item.Enabled = false;
            }
        }

        //Show All
        private void OptionBtn3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn3.Selected)
            {
                Load_Grid_All_Period();
                Button0.Item.Enabled = false;
                Button1.Item.Enabled = false;
            }
        }

        //Approve Button
        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (DocEntry_ODRF > 0)
            {
                SAPbobsCOM.ApprovalRequestsService oApprovalRequestsService = (SAPbobsCOM.ApprovalRequestsService)oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ApprovalRequestsService);
                SAPbobsCOM.ApprovalRequestsParams oApprovalRequestsParams = (SAPbobsCOM.ApprovalRequestsParams)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestsParams);
                SAPbobsCOM.ApprovalRequest oApprovalRequest = (SAPbobsCOM.ApprovalRequest)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequest);
                SAPbobsCOM.ApprovalRequestParams oApprovalRequestParams = (SAPbobsCOM.ApprovalRequestParams)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestParams);

                oApprovalRequestParams.Code = WddCode;
                
                //Approve request  
                oApprovalRequest = oApprovalRequestsService.GetApprovalRequest(oApprovalRequestParams);
                oApprovalRequest.ApprovalRequestDecisions.Add();
                oApprovalRequest.ApprovalRequestDecisions.Item(0).Remarks = EditText0.Value;
                oApprovalRequest.ApprovalRequestDecisions.Item(0).Status = SAPbobsCOM.BoApprovalRequestDecisionEnum.ardApproved;
                try
                {
                    oApprovalRequestsService.UpdateRequest(oApprovalRequest);
                    oApp.MessageBox("Phê duyệt thành công");
                    oApprovalRequest = oApprovalRequestsService.GetApprovalRequest(oApprovalRequestParams);
                    if (oApprovalRequest.Status == SAPbobsCOM.BoApprovalRequestStatusEnum.arsApproved)
                    {
                        try
                        {
                            //Approved Document Add to Valid Document
                            SAPbobsCOM.Documents oDraft = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                            oDraft.GetByKey(oApprovalRequest.ObjectEntry);
                            int ErrorCode = oDraft.SaveDraftToDocument();
                            if (ErrorCode == 0)
                                oApp.SetStatusBarMessage("Document added successfully", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                            else
                                oApp.SetStatusBarMessage(ErrorCode.ToString() + "|" + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                        }
                        catch
                        {
 
                        }
                    }
                }
                catch (Exception ex)
                {
                    oApp.MessageBox("Phê duyệt không thành công |" + ex.Message);
                }
            }

        }

        //Reject Button
        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (DocEntry_ODRF > 0)
            {
                SAPbobsCOM.ApprovalRequestsService oApprovalRequestsService = (SAPbobsCOM.ApprovalRequestsService)oCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ApprovalRequestsService);
                SAPbobsCOM.ApprovalRequestsParams oApprovalRequestsParams = (SAPbobsCOM.ApprovalRequestsParams)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestsParams);
                SAPbobsCOM.ApprovalRequest oApprovalRequest = (SAPbobsCOM.ApprovalRequest)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequest);
                SAPbobsCOM.ApprovalRequestParams oApprovalRequestParams = (SAPbobsCOM.ApprovalRequestParams)oApprovalRequestsService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestParams);

                oApprovalRequestParams.Code = WddCode;
                //oApprovalRequestsParams = oApprovalRequestsService.GetAllApprovalRequestsList();
                //oApprovalRequestParams = oApprovalRequestsParams.Item(oApprovalRequestsParams.Count - 1);

                //Approve request  
                oApprovalRequest = oApprovalRequestsService.GetApprovalRequest(oApprovalRequestParams);
                oApprovalRequest.ApprovalRequestDecisions.Add();
                oApprovalRequest.ApprovalRequestDecisions.Item(0).Remarks = EditText0.Value;
                oApprovalRequest.ApprovalRequestDecisions.Item(0).Status = SAPbobsCOM.BoApprovalRequestDecisionEnum.ardNotApproved;
                try
                {
                    oApprovalRequestsService.UpdateRequest(oApprovalRequest);
                    oApp.MessageBox("Từ chối phê duyệt thành công");
                }
                catch (Exception ex)
                {
                    oApp.MessageBox(ex.Message);
                }
            }

        }

        private void Grid0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            DocEntry_ODRF = 0;
            WddCode = 0;
            if (Grid0.Rows.SelectedRows.Count == 1)
            {
                int.TryParse(Grid0.DataTable.GetValue("DocEntry", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out DocEntry_ODRF);
                int.TryParse(Grid0.DataTable.GetValue("WddCode", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out WddCode);
                Load_Grid_Process();
                Load_Grid_Details();
            }


        }

    }
}
