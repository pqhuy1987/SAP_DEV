using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;

namespace U_ApproveJV
{
    [FormAttribute("U_ApproveJV.BILL_VP", "BILL_VP.b1f")]
    class BILL_VP : UserFormBase
    {
        public BILL_VP()
        {
            Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
        }
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        string BpName = "";
        string BpCode = "";
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        /// 
        private SAPbouiCOM.ComboBox ComboBox0;

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.StaticText StaticText7;

        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.CheckBox CheckBox0;

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.EditText EditText5;
        
        

        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Button Button4;

        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_3").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_13").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_2").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_btype").Specific));
            this.ComboBox1.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox1_ComboSelectAfter);
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("chk_new").Specific));
            this.CheckBox0.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox0_PressedAfter);
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_bpcode").Specific));
            this.EditText0.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.EditText0_LostFocusAfter);
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_prjn").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("txt_period").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("txt_fr").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("txt_to").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("grd_lst").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_n").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_v").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("bt_e").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("bt_d").Specific));
            this.Button3.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button3_PressedAfter);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("bt_l").Specific));
            this.Button4.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button4_PressedAfter);
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("txt_tu").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.CloseBefore += new CloseBeforeHandler(this.Form_CloseBefore);

        }

       

        private void OnCustomInitialize()
        {
            this.CheckBox0.ValOff = "0";
            this.CheckBox0.ValOn = "1";
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
            Load_Financial_Project();

        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            {
                SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
                oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                string sCFL_ID = null;
                sCFL_ID = oCFLEvento.ChooseFromListUID;
                SAPbouiCOM.Form oForm = null;
                oForm = Application.SBO_Application.Forms.Item(FormUID);
                SAPbouiCOM.ChooseFromList oCFL = null;
                oCFL = oForm.ChooseFromLists.Item(sCFL_ID);
                if (oCFLEvento.BeforeAction == false)
                {
                    SAPbouiCOM.DataTable oDataTable = null;
                    oDataTable = oCFLEvento.SelectedObjects;
                    BpCode = "";
                    try
                    {
                        BpCode = System.Convert.ToString(oDataTable.GetValue("CardCode", 0));
                    }
                    catch (Exception ex)
                    {

                    }
                    if (pVal.ItemUID == "txt_bpcode")
                    {
                        oForm.DataSources.UserDataSources.Item("UD_0").ValueEx = BpCode;
                    }

                }
            }

            if ((FormUID == "CFL1") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD))
            {
                System.Windows.Forms.Application.Exit();
            }
        }

        void Load_Financial_Project()
        {
            System.Data.DataTable tb_fprj = Get_List_FProject();
            if (tb_fprj.Rows.Count > 0)
            {
                foreach (DataRow r in tb_fprj.Rows)
                {
                    ComboBox0.ValidValues.Add(r["PrjCode"].ToString(), r["PrjName"].ToString());
                }
            }
        }

        System.Data.DataTable Get_List_FProject()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("VPBILL_GET_FPROJECT", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Username", oCompany.UserName);
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

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            try
            {
                EditText1.Value = ComboBox0.Selected.Description;
            }
            catch
            { }
        }

        private void CheckBox0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (CheckBox0.Checked)
            {
                EditText2.Item.Enabled = true;
                EditText3.Item.Enabled = true;
                EditText4.Item.Enabled = true;
                EditText5.Item.Enabled = true;
                ComboBox1.Item.Enabled = true;
                EditText2.Item.Click();
                try
                {
                    SAPbouiCOM.DataTable oDT = this.UIAPIRawForm.DataSources.DataTables.Item("DT_BILLVP");
                    if (oDT.Rows.Count > 0)
                    {
                        for (int i = 1; i <= oDT.Rows.Count; i++)
                        {
                            if (oDT.GetValue("Rejected", oDT.Rows.Count - i).ToString() == "N")
                            {
                                int tmp_period = 0;
                                int.TryParse(oDT.GetValue("Period", oDT.Rows.Count - i).ToString(), out tmp_period);
                                EditText2.Value = (tmp_period + 1).ToString();
                                DateTime tmp_dt = DateTime.Today;
                                DateTime.TryParse(oDT.GetValue("To", oDT.Rows.Count - 1).ToString(), out tmp_dt);
                                EditText3.Value = tmp_dt.AddDays(1).ToString("yyyyMMdd");
                                EditText4.Value = tmp_dt.AddDays(15).ToString("yyyyMMdd");
                                break;
                            }
                        }
                    }
                    else
                    {
                        EditText2.Value = "1";
                        EditText3.Value = DateTime.Today.AddDays(-14).ToString("yyyyMMdd");
                        EditText4.Value = DateTime.Today.ToString("yyyyMMdd");
                    }
                    Button0.Item.Enabled = true;
                }
                catch
                { }
                finally
                { }
            }
            else
            {
                EditText0.Item.Click();
                EditText2.Item.Enabled = false;
                EditText3.Item.Enabled = false;
                EditText4.Item.Enabled = false;
                EditText5.Item.Enabled = false;
                Button0.Item.Enabled = false;
            }

        }

        private void Load_Grid_Period(string pBP_Code, string pFProject)
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("VPBILL_GETLIST", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FProject", pFProject);
                cmd.Parameters.AddWithValue("@BP_Code", pBP_Code);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);
                this.UIAPIRawForm.Freeze(false);

                if (result.Rows.Count == 0)
                {
                    this.Button1.Item.Enabled = false;
                    this.Button2.Item.Enabled = false;
                }
                else
                {
                    this.Button1.Item.Enabled = true;
                    this.Button2.Item.Enabled = true;
                }

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

        private SAPbouiCOM.DataTable Convert_SAP_DataTable(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_BILLVP"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_BILLVP");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_BILLVP");
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
                    oDT.SetValue(c.ColumnName, oDT.Rows.Count - 1, r[c.ColumnName]);
                }
            }

            return oDT;
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

        System.Data.DataTable Get_Data_BILLVP(string pFinancialProject, DateTime pTo_Date, string pBPCode)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("VPBILL_GETDATA", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                cmd.Parameters.AddWithValue("@To_Date", pTo_Date);
                cmd.Parameters.AddWithValue("@BP_Code", pBPCode);
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

        System.Data.DataTable Get_Approve_Process_BILLVP()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("VPBILL_GetList_Approve", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Usr", oCompany.UserName);
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

        private void ComboBox1_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (ComboBox1.Selected.Value == "1")
                EditText5.Item.Enabled = true;
            else
                EditText5.Item.Enabled = false;

        }

        private void EditText0_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (!String.IsNullOrEmpty(EditText0.Value))
            {
                BpCode = EditText0.Value;
                if (!string.IsNullOrEmpty(ComboBox0.Value))
                    Load_Grid_Period(EditText0.Value, ComboBox0.Selected.Value);
            }
        }
        
        private void Get_BpName()
        {
            try
            {
                SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oR_RecordSet.DoQuery(string.Format("Select CardName from ORCD where CardCode='{0}'", BpCode.Replace(';', ' ')));
                if (oR_RecordSet.RecordCount > 0)
                {
                    BpName = oR_RecordSet.Fields.Item("CardName").Value.ToString();
                }
                else
                {
                    BpName = "";
                }
            }
            catch
            {

            }
        }
        
        //Add Button
        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService sCmp = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralData oChild = null;
            SAPbobsCOM.GeneralDataCollection oChildren = null;
            try
            {
                BpCode = EditText0.Value;
                Get_BpName();

                sCmp = oCompany.GetCompanyService();
                oGeneralService = sCmp.GetGeneralService("BILLVP");
                //Create data for new row in main UDO
                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                //Financial Project
                oGeneralData.SetProperty("U_FProject", ComboBox0.Selected.Value);
                //Project Name
                oGeneralData.SetProperty("U_ProjectName", EditText1.Value);
                //BpCode
                oGeneralData.SetProperty("U_BPCode", BpCode);
                //BpName
                oGeneralData.SetProperty("U_BPName", BpName);
                //Bill Type
                int btype = 2;
                int.TryParse(ComboBox1.Selected.Value, out btype);
                oGeneralData.SetProperty("U_BType", btype);
                //Date From
                oGeneralData.SetProperty("U_DateFr", DateTime.ParseExact(EditText3.Value, "yyyyMMdd", CultureInfo.InvariantCulture));
                //Date To
                oGeneralData.SetProperty("U_DateTo", DateTime.ParseExact(EditText4.Value, "yyyyMMdd", CultureInfo.InvariantCulture));
                //Period
                oGeneralData.SetProperty("U_Period", EditText2.Value);
                //Add Bill Tam ung
                if (ComboBox1.Selected.Value == "1")
                {
                    Double tmp_tu = 0;
                    Double.TryParse(EditText5.Value, out tmp_tu);
                    oGeneralData.SetProperty("U_Tamung", tmp_tu);
                }
                //Add Bill thanh toan
                else if (ComboBox1.Selected.Value == "2")
                {
                    DataTable rs = Get_Data_BILLVP(ComboBox0.Selected.Value, DateTime.ParseExact(EditText4.Value, "yyyyMMdd", CultureInfo.InvariantCulture), BpCode);
                    Double t = 0;
                    if (rs.Rows.Count > 0)
                    {
                        //Create data for a row in the child table
                        oChildren = oGeneralData.Child("BILLVP1");
                        foreach (DataRow r in rs.Rows)
                        {
                            oChild = oChildren.Add();
                            //So chung tu
                            oChild.SetProperty("U_GRPO_Key", int.Parse(r["GRPOKey"].ToString()));
                            //Line details
                            oChild.SetProperty("U_GRPO_Line", int.Parse(r["GRPORowKey"].ToString()));
                            //Ma phong ban
                            oChild.SetProperty("U_DistRule", r["MaPB"].ToString());
                            //Ten phong ban
                            oChild.SetProperty("U_DisRule_Name", r["TenPB"].ToString());
                            //Du an
                            oChild.SetProperty("U_Project", r["DA"].ToString());
                            //Level 1
                            oChild.SetProperty("U_Level1", r["U_ParentID1"].ToString());
                            //Level 1 Name
                            oChild.SetProperty("U_Level1Name", r["Name1"].ToString());
                            //Level 2
                            oChild.SetProperty("U_Level2", r["U_ParentID2"].ToString());
                            //Level 2 Name
                            oChild.SetProperty("U_Level2Name", r["Name2"].ToString());
                            //Level 3
                            oChild.SetProperty("U_Level3", r["U_ParentID3"].ToString());
                            //Level 3 Name
                            oChild.SetProperty("U_Level3Name", r["Name3"].ToString());
                            //Level 4
                            oChild.SetProperty("U_Level4", r["U_ParentID4"].ToString());
                            //Level 4 Name
                            oChild.SetProperty("U_Level4Name", r["Name4"].ToString());
                            //Level 5
                            oChild.SetProperty("U_Level5", r["U_ParentID5"].ToString());
                            //Level 5 Name
                            oChild.SetProperty("U_Level5Name", r["Name5"].ToString());
                            //Ma CP
                            oChild.SetProperty("U_MaCP", r["MaCP"].ToString());
                            //Ten CP
                            oChild.SetProperty("U_TenCP", r["TenCP"].ToString());
                            //Noi dung
                            oChild.SetProperty("U_Noidung", r["DetailsName"].ToString());
                            //GT
                            t = 0;
                            Double.TryParse(r["Gross_Total"].ToString(),out t);
                            oChild.SetProperty("U_GrossTotal", t);
                            //GT no VAT
                            t = 0;
                            Double.TryParse(r["Total"].ToString(), out t);
                            oChild.SetProperty("U_Total", t);
                        }
                    }
                }
                //Add Quy trinh duyet
                DataTable rs2 = Get_Approve_Process_BILLVP();
                if (rs2.Rows.Count > 0)
                {
                    oChildren = oGeneralData.Child("BILLVP2");
                    foreach (DataRow r in rs2.Rows)
                    {
                        oChild = oChildren.Add();
                        oChild.SetProperty("U_Level", r["LEVEL"].ToString());
                        oChild.SetProperty("U_Position", r["Position"].ToString());
                        oChild.SetProperty("U_DeptName", r["DeptName"].ToString());
                        oChild.SetProperty("U_PosName", r["PosName"].ToString());
                    }
                }
                //Add UDO
                oGeneralParams = oGeneralService.Add(oGeneralData);
                if (!string.IsNullOrEmpty(oGeneralParams.GetProperty("DocEntry").ToString()))
                {
                    oApp.SetStatusBarMessage("Add Completed - DocEntry: " + oGeneralParams.GetProperty("DocEntry").ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    Load_Grid_Period(BpCode, ComboBox0.Selected.Value);

                }
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message, 1, "OK");
            }

        }

        //View Button
        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (Grid0.Rows.SelectedRows.Count == 1)
                {
                    oApp.ActivateMenuItem("BILLVP");
                    SAPbouiCOM.Form frm = Application.SBO_Application.Forms.ActiveForm;
                    frm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    frm.Items.Item("1_U_E").Enabled = true;
                    //find Selected Key Matrix
                    string DocNum = Grid0.DataTable.GetValue("Document Number", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();

                    ((SAPbouiCOM.EditText)frm.Items.Item("1_U_E").Specific).Value = DocNum;
                    frm.Items.Item("1").Click();
                    frm.EnableMenu("1282", false);  // Add New Record
                    frm.EnableMenu("1288", false);  // Next Record
                    frm.EnableMenu("1289", false);  // Pevious Record
                    frm.EnableMenu("1290", false);  // First Record
                    frm.EnableMenu("1291", false);  // Last record
                    frm.EnableMenu("1283", false);  // Remove
                    frm.EnableMenu("1284", false);  // Cancel
                    frm.EnableMenu("1286", false);  // Close
                    frm.EnableMenu("1304", false);  //Refresh
                    frm.EnableMenu("BILLVP_Remove_Line", false);
                    frm.EnableMenu("BILLVP_Add_Line", false);
                    //Disable Header
                    foreach (SAPbouiCOM.Item it in frm.Items)
                    {
                        if (it.Type == SAPbouiCOM.BoFormItemTypes.it_EDIT || it.Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                        {
                            switch (it.Description.Trim())
                            {
                                case "Financial Project":
                                    it.Enabled = false;
                                    break;
                                case "Project Name":
                                    it.Enabled = false;
                                    break;
                                case "BP Code":
                                    it.Enabled = false;
                                    break;
                                case "BP Name":
                                    it.Enabled = false;
                                    break;
                                case "Bill Type":
                                    it.Enabled = false;
                                    break;
                                case "Period":
                                    it.Enabled = false;
                                    break;
                                case "From Date":
                                    it.Enabled = false;
                                    break;
                                case "To Date":
                                    it.Enabled = false;
                                    break;
                                case "Tạm ứng":
                                    it.Enabled = false;
                                    break;
                            }
                        }
                    }
                }
                else
                {
                    oApp.MessageBox("Please select record !");
                }
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }

        }

        //Edit Button
        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            
            try
            {
                if (Grid0.Rows.SelectedRows.Count == 1)
                {
                    oApp.ActivateMenuItem("BILLVP");
                    SAPbouiCOM.Form frm = Application.SBO_Application.Forms.ActiveForm;
                    frm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    frm.Items.Item("1_U_E").Enabled = true;
                    //find Selected Key Matrix
                    string DocNum = Grid0.DataTable.GetValue("Document Number", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();

                    ((SAPbouiCOM.EditText)frm.Items.Item("1_U_E").Specific).Value = DocNum;
                    frm.Items.Item("1").Click();
                    frm.EnableMenu("1282", false);  // Add New Record
                    frm.EnableMenu("1288", false);  // Next Record
                    frm.EnableMenu("1289", false);  // Pevious Record
                    frm.EnableMenu("1290", false);  // First Record
                    frm.EnableMenu("1291", false);  // Last record
                    frm.EnableMenu("1283", false);  // Remove
                    frm.EnableMenu("1284", false);  // Cancel
                    frm.EnableMenu("1286", false);  // Close
                    frm.EnableMenu("1304", false);  //Refresh
                    frm.EnableMenu("BILLVP_Remove_Line", false);
                    frm.EnableMenu("BILLVP_Add_Line", false);
                    //Disable Header
                    foreach (SAPbouiCOM.Item it in frm.Items)
                    {
                        if (it.Type == SAPbouiCOM.BoFormItemTypes.it_EDIT || it.Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                        {
                            switch (it.Description.Trim())
                            {
                                case "Financial Project":
                                    it.Enabled = true;
                                    break;
                                case "Project Name":
                                    it.Enabled = true;
                                    break;
                                case "BP Code":
                                    it.Enabled = true;
                                    break;
                                case "BP Name":
                                    it.Enabled = true;
                                    break;
                                case "Bill Type":
                                    it.Enabled = true;
                                    break;
                                case "Period":
                                    it.Enabled = true;
                                    break;
                                case "From Date":
                                    it.Enabled = true;
                                    break;
                                case "To Date":
                                    it.Enabled = true;
                                    break;
                                case "Tạm ứng":
                                    it.Enabled = true;
                                    break;
                            }
                        }
                    }
                }
                else
                {
                    oApp.MessageBox("Please select record !");
                }
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }

        }

        //Delete Button
        private void Button3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService sCmp = null;
            try
            {
                if (Grid0.Rows.SelectedRows.Count == 1)
                {
                    if (oApp.MessageBox("Are you sure you want to delete this record ?", 2, "Yes", "No") == 1)
                    {
                        int DocEntry = 0;
                        string t = Grid0.DataTable.GetValue("Document Number", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                        int.TryParse(t, out DocEntry);

                        sCmp = oCompany.GetCompanyService();
                        oGeneralService = sCmp.GetGeneralService("BILLVP");
                        oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oGeneralParams.SetProperty("DocEntry", DocEntry);
                        oGeneralService.Delete(oGeneralParams);
                        oApp.SetStatusBarMessage("Delete Complete", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        Load_Grid_Period(EditText0.Value, ComboBox0.Selected.Value);
                    }
                }
                else
                {
                    oApp.MessageBox("Please select record !");
                }
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
        }

        //Load Button
        private void Button4_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (!string.IsNullOrEmpty(EditText1.Value))
            {
                if (!string.IsNullOrEmpty(ComboBox0.Value))
                {
                    Load_Grid_Period(EditText0.Value, ComboBox0.Selected.Value);
                }
            }
        }

        private void Form_CloseBefore(SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                oCompany.Disconnect();
            }
            catch
            { }
            //throw new System.NotImplementedException();

        }

    }
}
