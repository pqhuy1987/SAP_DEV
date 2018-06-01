using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Xml;
namespace U_KLTT
{
    [FormAttribute("U_KLTT.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
            Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            Application.SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
        }
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Button Button4;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.ComboBox ComboBox2;
        private SAPbouiCOM.ComboBox ComboBox3;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.Button Button6;
        private SAPbouiCOM.CheckBox CheckBox0;
        private SAPbouiCOM.Grid Grid0;
        
        //private SAPbouiCOM.StaticText StaticText0;
        //private SAPbouiCOM.StaticText StaticText1;
        //private SAPbouiCOM.StaticText StaticText2;
        //private SAPbouiCOM.StaticText StaticText4;
        //private SAPbouiCOM.StaticText StaticText5;
        //private SAPbouiCOM.StaticText StaticText6;

        public override void OnInitializeComponent()
        {
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cb_fipro").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_grpb").Specific));
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_billt").Specific));
            this.ComboBox3 = ((SAPbouiCOM.ComboBox)(this.GetItem("cbo_put").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_pname").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_bp").Specific));
            this.EditText1.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.EditText1_LostFocusAfter);
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("txt_period").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("txt_fr").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("txt_to").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("gr_kybill").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_add").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_view").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("bt_update").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("chk_new").Specific));
            this.CheckBox0.PressedAfter += new SAPbouiCOM._ICheckBoxEvents_PressedAfterEventHandler(this.CheckBox0_PressedAfter);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("bt_load").Specific));
            this.Button3.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button3_PressedAfter);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("bt_report").Specific));
            this.Button4.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button4_PressedAfter);
            this.Button6 = ((SAPbouiCOM.Button)(this.GetItem("bt_del").Specific));
            this.Button6.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button6_PressedAfter);
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
            // events handled by SBO_Application_ItemEvent
            // += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent); 
            Load_Financial_Project();
        }

        System.Data.DataTable Get_Data_KLTT(string pFinancialProject, string pType, DateTime pTo_Date, string pBPCode, string pBGroup, string pPurchaseType)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_GETDATA", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                cmd.Parameters.AddWithValue("@To_Date", pTo_Date);
                cmd.Parameters.AddWithValue("@BP_Code", pBPCode);
                cmd.Parameters.AddWithValue("@Type", pType);
                cmd.Parameters.AddWithValue("@BGroup", pBGroup);
                cmd.Parameters.AddWithValue("@PurchaseType", pPurchaseType);
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

        System.Data.DataTable Get_Data_KLTT_NT(string pFinancialProject, string pType, int pPeriod, string pBPCode, string pBGroup, string pPurchaseType)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_GETDATA_NHANTRI", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                cmd.Parameters.AddWithValue("@Period", pPeriod);
                cmd.Parameters.AddWithValue("@BP_Code", pBPCode);
                cmd.Parameters.AddWithValue("@Type", pType);
                cmd.Parameters.AddWithValue("@BGroup", pBGroup);
                cmd.Parameters.AddWithValue("@PurchaseType", pPurchaseType);
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
        
        System.Data.DataTable GetList_Approve(string pBGroup,string pNhanTri = "N")
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_GETLIST_APPROVE", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@BGroup", pBGroup);
                cmd.Parameters.AddWithValue("@Nhantri", pNhanTri);
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

        System.Data.DataTable Load_Data_KLTT_AI(string pFinancialProject, int pPeriod, string pBPCode, DateTime pToDate, string pCGroup = "", string pPUTYPE = "")
        {
            DataTable result = new DataTable();

            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_GET_ADDITIONALINFO", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                cmd.Parameters.AddWithValue("@Period", pPeriod);
                cmd.Parameters.AddWithValue("@BP_Code", pBPCode);
                cmd.Parameters.AddWithValue("@CGroup", pCGroup);
                cmd.Parameters.AddWithValue("@PurchaseType", pPUTYPE);
                cmd.Parameters.AddWithValue("@ToDate", pToDate);
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
        
        System.Data.DataTable Load_Data_KLTT(string pFinancialProject, string pType, int pPeriod, string pBPCode, string pCGroup="", int pDocEntry = 0, string pPUTYPE = "")
        {
            DataTable result = new DataTable();

            SqlCommand cmd = null;
            try
            {
                if(pType == "SPP")
                {
                    cmd = new SqlCommand("KLTT_TOTAL", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                    cmd.Parameters.AddWithValue("@Period", pPeriod);
                    cmd.Parameters.AddWithValue("@BP_Code", pBPCode);
                    cmd.Parameters.AddWithValue("@PUType", pPUTYPE);
                    cmd.Parameters.AddWithValue("@BGroup", pCGroup);
                }
                else
                {
                    cmd = new SqlCommand("KLTT_LOADDATA", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                    cmd.Parameters.AddWithValue("@Period", pPeriod);
                    cmd.Parameters.AddWithValue("@BP_Code", pBPCode);
                    cmd.Parameters.AddWithValue("@Type", pType);
                    cmd.Parameters.AddWithValue("@DocEntry", pDocEntry);
                }
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

        System.Data.DataTable Get_List_Sub_Level(int pKLTT_DocEntry, int pParentID, int pLevel, string pType)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {

                cmd = new SqlCommand("KLTT_Get_List_Sub_Level", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@KLTT_DocEntry", pKLTT_DocEntry);
                cmd.Parameters.AddWithValue("@Parent_ID", pParentID);
                cmd.Parameters.AddWithValue("@Level", pLevel);
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
                cmd = new SqlCommand("KLTT_GET_FPROJECT", conn);
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

        void Load_Info()
        {
            try
            {
                EditText0.Value = ComboBox0.Selected.Description;
            }
            catch
            {}

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
            SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_KLTTList"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_KLTTList");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_KLTTList");
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
                {}

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

        private void Load_Grid_Period(string pBP_Code, string pFProject, string pBGroup, string pPurchaseType)
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_GETLIST", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FProject", pFProject);
                cmd.Parameters.AddWithValue("@BP_Code", pBP_Code);
                cmd.Parameters.AddWithValue("@BGroup", pBGroup);
                cmd.Parameters.AddWithValue("@PurchaseType", pPurchaseType);
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

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Load_Info();
        }

        private void CheckBox0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            
            if (this.CheckBox0.Checked)
            {
                this.EditText2.Item.Enabled = true;
                this.EditText3.Item.Enabled = true;
                this.EditText4.Item.Enabled = true;
                this.ComboBox2.Item.Enabled = true;
                EditText2.Item.Click();
                try
                {
                    SAPbouiCOM.DataTable oDT = this.UIAPIRawForm.DataSources.DataTables.Item("DT_KLTTList");
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
                this.EditText2.Item.Enabled = false;
                this.EditText3.Item.Enabled = false;
                this.EditText4.Item.Enabled = false;
                this.ComboBox2.Item.Enabled = false;
                Button0.Item.Enabled = false;
            }
        }

        private void EditText1_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (!String.IsNullOrEmpty(EditText1.Value))
            {
                if (!string.IsNullOrEmpty(ComboBox0.Value))
                    Load_Grid_Period(this.EditText1.Value.Trim(), ComboBox0.Selected.Value, ComboBox1.Selected.Value,ComboBox3.Selected.Value);
            }
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
                    string val = null;
                    try
                    {
                        val = System.Convert.ToString(oDataTable.GetValue(0, 0));
                    }
                    catch (Exception ex)
                    {

                    }
                    if ((pVal.ItemUID == "txt_bp") | (pVal.ItemUID == "txt_bp"))
                    {
                        oForm.DataSources.UserDataSources.Item("UD_2").ValueEx = val;
                    }

                }
            }
            else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK && pVal.ItemUID == "gr_kybill" && pVal.FormUID == "frm_kltt" && pVal.Action_Success == true)
            {
                if (Grid0.Rows.SelectedRows.Count == 1)
                {
                    try
                    {
                        oApp.ActivateMenuItem("KLTT");
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
                        frm.EnableMenu("KLTT_Remove_Line", false);
                        frm.EnableMenu("KLTT_Add_Line", false);
                        frm.Items.Item("0_U_G").Enabled = false;
                        frm.Items.Item("1_U_G").Enabled = false;
                        frm.Items.Item("2_U_G").Enabled = false;
                        frm.Items.Item("3_U_G").Enabled = false;
                        frm.Items.Item("4_U_G").Enabled = false;
                        frm.Items.Item("5_U_G").Enabled = false;
                        frm.Items.Item("6_U_G").Enabled = false;
                        //Disable Header
                        frm.Items.Item("20_U_E").Enabled = false;
                        frm.Items.Item("21_U_E").Enabled = false;
                        frm.Items.Item("22_U_E").Enabled = false;
                        frm.Items.Item("23_U_E").Enabled = false;
                        frm.Items.Item("24_U_E").Enabled = false;
                        frm.Items.Item("25_U_E").Enabled = false;
                        frm.Items.Item("26_U_E").Enabled = false;
                        frm.Items.Item("27_U_E").Enabled = false;
                        //Disable all column Matrix
                        Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("0_U_G").Specific, false);
                        Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("1_U_G").Specific, false);
                        Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("2_U_G").Specific, false);
                        Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("3_U_G").Specific, false);
                        Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("4_U_G").Specific, false);
                        Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("5_U_G").Specific, false);
                        Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("6_U_G").Specific, false);
                        Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("7_U_G").Specific, false);
                    }
                    catch
                    {
 
                    }
                }
            }

            if ((FormUID == "CFL1") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD))
            {
                System.Windows.Forms.Application.Exit();
            }
        }

        private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.BeforeAction && pVal.FormUID.StartsWith("UDO_F_KLTT"))
            {
                try
                {
                    pVal.RemoveFromContent("KLTT_Add_Line");
                    pVal.RemoveFromContent("KLTT_Remove_Line");
                }
                catch
                { }
            }
        }

        private void Editable_Column_Matrix(SAPbouiCOM.Matrix pMatrixItem, bool editable)
        {
            try
            {
                foreach (SAPbouiCOM.Column c in pMatrixItem.Columns)
                {
                    c.Editable = editable;
                }
            }
            catch
            { }
        }

        private void Editable_Column_Matrix_EditMode(SAPbouiCOM.Matrix pMatrixItem)
        {
            try
            {
                foreach (SAPbouiCOM.Column c in pMatrixItem.Columns)
                {
                    if (c.Title != "Phần trăm hoàn thành" && c.Title != "Giá trị thực hiện")
                        c.Editable = false;
                    else
                        c.Editable = true;
                }
            }
            catch
            { }
        }

        public String convert(decimal total)
        {
            try
            {
                string rs = "";
                total = Math.Abs(Math.Round(total, 0));
                string[] ch = { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };
                string[] rch = { "lẻ", "mốt", "", "", "", "lăm" };
                string[] u = { "", "mươi", "trăm", "ngàn", "", "", "triệu", "", "", "tỷ", "", "", "ngàn", "", "", "triệu" };
                string nstr = total.ToString();
                int[] n = new int[nstr.Length];
                int len = n.Length;
                for (int i = 0; i < len; i++)
                {
                    n[len - 1 - i] = Convert.ToInt32(nstr.Substring(i, 1));
                }
                for (int i = len - 1; i >= 0; i--)
                {
                    if (i % 3 == 2)// số 0 ở hàng trăm
                    {
                        if (n[i] == 0 && n[i - 1] == 0 && n[i - 2] == 0) continue;//nếu cả 3 số là 0 thì bỏ qua không đọc
                    }
                    else if (i % 3 == 1) // số ở hàng chục
                    {
                        if (n[i] == 0)
                        {
                            if (n[i - 1] == 0) { continue; }// nếu hàng chục và hàng đơn vị đều là 0 thì bỏ qua.
                            else
                            {
                                rs += " " + rch[n[i]]; continue;// hàng chục là 0 thì bỏ qua, đọc số hàng đơn vị
                            }
                        }
                        if (n[i] == 1)//nếu số hàng chục là 1 thì đọc là mười
                        {
                            rs += " mười"; continue;
                        }
                    }
                    else if (i != len - 1)// số ở hàng đơn vị (không phải là số đầu tiên)
                    {
                        if (n[i] == 0)// số hàng đơn vị là 0 thì chỉ đọc đơn vị
                        {
                            if (i + 2 <= len - 1 && n[i + 2] == 0 && n[i + 1] == 0) continue;
                            rs += " " + (i % 3 == 0 ? u[i] : u[i % 3]);
                            continue;
                        }
                        if (n[i] == 1)// nếu là 1 thì tùy vào số hàng chục mà đọc: 0,1: một / còn lại: mốt
                        {
                            rs += " " + ((n[i + 1] == 1 || n[i + 1] == 0) ? ch[n[i]] : rch[n[i]]);
                            rs += " " + (i % 3 == 0 ? u[i] : u[i % 3]);
                            continue;
                        }
                        if (n[i] == 5) // cách đọc số 5
                        {
                            if (n[i + 1] != 0) //nếu số hàng chục khác 0 thì đọc số 5 là lăm
                            {
                                rs += " " + rch[n[i]];// đọc số 
                                rs += " " + (i % 3 == 0 ? u[i] : u[i % 3]);// đọc đơn vị
                                continue;
                            }
                        }
                    }
                    rs += (rs == "" ? " " : ", ") + ch[n[i]];// đọc số
                    rs += " " + (i % 3 == 0 ? u[i] : u[i % 3]);// đọc đơn vị
                }
                if (rs[rs.Length - 1] != ' ')
                    rs += " đồng.";
                else
                    rs += "đồng.";
                if (rs.Length > 2)
                {
                    string rs1 = rs.Substring(0, 2);
                    rs1 = rs1.ToUpper();
                    rs = rs.Substring(2);
                    rs = rs1 + rs;
                }
                return rs.Trim().Replace("lẻ,", "lẻ").Replace("mươi,", "mươi").Replace("trăm,", "trăm").Replace("mười,", "mười");
            }
            catch
            {
                return "Error";
            }
        }

        private bool Check_NT(string pBpCode)
        {
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(string.Format("Select GroupCode from OCRD where CardCode='{0}'", pBpCode));
            if (oR_RecordSet.RecordCount > 0)
            {
                string BPGCode = oR_RecordSet.Fields.Item("GroupCode").Value.ToString();
                if (BPGCode=="112") return true;
                else return false;
            }
            return false;
        }

        //Add New
        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = null;
            if (ComboBox2.Value == "")
            {
                oApp.MessageBox("Please select Bill Type !");
                return;
            }
            if (ComboBox3.Value == "")
            {
                oApp.MessageBox("Please select Purchase Type !");
                return;
            }
            try
            {
                string fProject = ComboBox0.Selected.Value;
                //string ProjectKey = ComboBox0.Selected.Value;
                string BType = ComboBox2.Selected.Value;
                string BGroup = ComboBox1.Selected.Value;
                string PurchaseType = ComboBox3.Selected.Value;
                string BPCode = EditText1.Value;
                string QLNT = "";
                int Period = 1;
                int.TryParse(EditText2.Value, out Period);
                bool NT = Check_NT(BPCode);
                DateTime Fr_Date = DateTime.ParseExact(EditText3.Value, "yyyyMMdd", CultureInfo.InvariantCulture);
                DateTime To_Date = DateTime.ParseExact(EditText4.Value, "yyyyMMdd", CultureInfo.InvariantCulture);

                Application.SBO_Application.ActivateMenuItem("KLTT");
                oForm = Application.SBO_Application.Forms.ActiveForm;
                //Fill Header

                ((SAPbouiCOM.EditText)oForm.Items.Item("20_U_E").Specific).Value = fProject; //Financial Project
                ((SAPbouiCOM.EditText)oForm.Items.Item("21_U_E").Specific).Value = Fr_Date.ToString("yyyyMMdd"); //Date From
                ((SAPbouiCOM.EditText)oForm.Items.Item("22_U_E").Specific).Value = To_Date.ToString("yyyyMMdd"); //Date To
                ((SAPbouiCOM.EditText)oForm.Items.Item("24_U_E").Specific).Value = BPCode; //Bp Code
                ((SAPbouiCOM.EditText)oForm.Items.Item("25_U_E").Specific).Value = Period.ToString(); //Period
                ((SAPbouiCOM.EditText)oForm.Items.Item("26_U_E").Specific).Value = DateTime.Now.ToString("yyyyMMdd"); //Created Date
                ((SAPbouiCOM.ComboBox)oForm.Items.Item("29_U_Cb").Specific).Select(BGroup, SAPbouiCOM.BoSearchKey.psk_ByValue); //Bill Group
                ((SAPbouiCOM.ComboBox)oForm.Items.Item("30_U_Cb").Specific).Select(BType, SAPbouiCOM.BoSearchKey.psk_ByValue); // Bill Type
                ((SAPbouiCOM.ComboBox)oForm.Items.Item("32_U_Cb").Specific).Select(PurchaseType, SAPbouiCOM.BoSearchKey.psk_ByValue); //Purchase Type
                
                //Check Thanh toan tam ung

                DataTable Additional_Info = Load_Data_KLTT_AI(fProject, Period, BPCode, To_Date, BGroup, PurchaseType);
                if (Additional_Info.Rows.Count > 0)
                {
                    //Get Max Tam ung
                    decimal max_tu = 0;
                    decimal.TryParse(Additional_Info.Rows[0]["GTTU"].ToString(), out max_tu);
                    //Thuoc doi NT
                    QLNT = Additional_Info.Rows[0]["CTQLDTC"].ToString();
                    if (!string.IsNullOrEmpty(QLNT) && QLNT != "")
                        ((SAPbouiCOM.EditText)oForm.Items.Item("33_U_E").Specific).Value = QLNT;
                    //BILL NHAN TRI
                    if (BType == "1")
                    {
                        //Bill Tam ung
                        //Gia tri tam ung
                        //((SAPbouiCOM.EditText)oForm.Items.Item("32_U_E").Specific).Value = max_tu.ToString();
                        //Fill Value I - Phe duyet
                        oForm.PaneLevel = 9;
                        oForm.Freeze(true);
                        DataTable APPRV1 = GetList_Approve(BGroup);
                        if (APPRV1.Rows.Count > 0)
                        {
                            SAPbouiCOM.Matrix oMtx1 = (SAPbouiCOM.Matrix)oForm.Items.Item("8_U_G").Specific;
                            for (int i = 0; i < APPRV1.Rows.Count; i++)
                            {
                                DataRow r = APPRV1.Rows[i];
                                ((SAPbouiCOM.ComboBox)oMtx1.Columns.Item("C_8_2").Cells.Item(i + 1).Specific).Select(r["LEVEL"].ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                                ((SAPbouiCOM.ComboBox)oMtx1.Columns.Item("C_8_5").Cells.Item(i + 1).Specific).Select(r["Position"].ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                                if (i < APPRV1.Rows.Count - 1)
                                {
                                    oMtx1.AddRow();
                                    ((SAPbouiCOM.EditText)oMtx1.Columns.Item("C_8_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                                }
                            }
                        }
                        oForm.Freeze(false);

                        oApp.StatusBar.SetText("Load Data Completed", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        return;
                    }
                }
                else
                {
                    oApp.StatusBar.SetText("Missing Contract :(", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                }

                //Fill Value A - Theo hop dong
                DataTable A = new DataTable();
                if (NT)
                    A = Get_Data_KLTT_NT(fProject, "A", Period, BPCode, BGroup, PurchaseType);
                else
                    A = Get_Data_KLTT(fProject, "A", To_Date, BPCode, BGroup, PurchaseType);
                oForm.PaneLevel = 1;
                oForm.Freeze(true);
                SAPbouiCOM.Matrix oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("0_U_G").Specific;
                if (A.Rows.Count > 0)
                {
                    for (int i = 0; i < A.Rows.Count; i++)
                    {
                        DataRow r = A.Rows[i];
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_2").Cells.Item(i + 1).Specific).Value = r["GoiThauKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_3").Cells.Item(i + 1).Specific).Value = r["GoiThauName"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_4").Cells.Item(i + 1).Specific).Value = r["SubProjectKey"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_5").Cells.Item(i + 1).Specific).Value = r["SubProjectName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_6").Cells.Item(i + 1).Specific).Value = r["GRPOKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_7").Cells.Item(i + 1).Specific).Value = r["GRPORowKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_8").Cells.Item(i + 1).Specific).Value = r["DetailsName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_9").Cells.Item(i + 1).Specific).Value = r["DetailsWork"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_10").Cells.Item(i + 1).Specific).Value = r["U_ParentID1"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_11").Cells.Item(i + 1).Specific).Value = r["Name1"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_12").Cells.Item(i + 1).Specific).Value = r["U_ParentID2"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_13").Cells.Item(i + 1).Specific).Value = r["Name2"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_14").Cells.Item(i + 1).Specific).Value = r["U_ParentID3"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_15").Cells.Item(i + 1).Specific).Value = r["Name3"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_16").Cells.Item(i + 1).Specific).Value = r["U_ParentID4"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_17").Cells.Item(i + 1).Specific).Value = r["Name4"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_18").Cells.Item(i + 1).Specific).Value = r["U_ParentID5"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_19").Cells.Item(i + 1).Specific).Value = r["Name5"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_20").Cells.Item(i + 1).Specific).Value = r["TYPE"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_21").Cells.Item(i + 1).Specific).Value = r["UoM"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_22").Cells.Item(i + 1).Specific).Value = r["UPrice"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_23").Cells.Item(i + 1).Specific).Value = r["Quantity"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_24").Cells.Item(i + 1).Specific).Value = r["Total"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_25").Cells.Item(i + 1).Specific).Value = BType == "3" ? "100" : r["Last_Complete_Rate"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_26").Cells.Item(i + 1).Specific).Value = BType == "3" ? r["Total"].ToString() : r["Last_Complete_Amount"].ToString();
                        if (i < A.Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }
                oForm.Items.Item("20_U_E").Click();
                //Invisible Column 
                //oMtx.Columns.Item("C_0_2").Visible = false;
                //oMtx.Columns.Item("C_0_4").Visible = false;
                //oMtx.Columns.Item("C_0_6").Visible = false;
                //oMtx.Columns.Item("C_0_7").Visible = false;
                //Disable All Column except 2 last column
                oMtx.AutoResizeColumns();
                foreach (SAPbouiCOM.Column c in oMtx.Columns)
                {
                    if (c.Title != "Phần trăm hoàn thành" && c.Title != "Giá trị thực hiện")
                        c.Editable = false;
                    else
                        c.Editable = true;
                }
                oForm.Freeze(false);

                //Fill Value B - Phat sinh
                DataTable B = new DataTable();// Get_Data_KLTT(fProject, "B", To_Date, BPCode, BGroup, PurchaseType);
                if (NT)
                    B = Get_Data_KLTT_NT(fProject, "B", Period, BPCode, BGroup, PurchaseType);
                else
                    B = Get_Data_KLTT(fProject, "B", To_Date, BPCode, BGroup, PurchaseType);
                oForm.PaneLevel = 2;
                oForm.Freeze(true);
                oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("1_U_G").Specific;
                if (B.Rows.Count > 0)
                {
                    for (int i = 0; i < B.Rows.Count; i++)
                    {
                        DataRow r = B.Rows[i];
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_2").Cells.Item(i + 1).Specific).Value = r["GoiThauKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_3").Cells.Item(i + 1).Specific).Value = r["GoiThauName"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_4").Cells.Item(i + 1).Specific).Value = r["SubProjectKey"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_0_5").Cells.Item(i + 1).Specific).Value = r["SubProjectName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_6").Cells.Item(i + 1).Specific).Value = r["GRPOKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_7").Cells.Item(i + 1).Specific).Value = r["GRPORowKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_8").Cells.Item(i + 1).Specific).Value = r["DetailsName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_9").Cells.Item(i + 1).Specific).Value = r["DetailsWork"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_10").Cells.Item(i + 1).Specific).Value = r["U_ParentID1"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_11").Cells.Item(i + 1).Specific).Value = r["Name1"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_12").Cells.Item(i + 1).Specific).Value = r["U_ParentID2"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_13").Cells.Item(i + 1).Specific).Value = r["Name2"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_14").Cells.Item(i + 1).Specific).Value = r["U_ParentID3"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_15").Cells.Item(i + 1).Specific).Value = r["Name3"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_16").Cells.Item(i + 1).Specific).Value = r["U_ParentID4"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_17").Cells.Item(i + 1).Specific).Value = r["Name4"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_18").Cells.Item(i + 1).Specific).Value = r["U_ParentID5"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_19").Cells.Item(i + 1).Specific).Value = r["Name5"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_20").Cells.Item(i + 1).Specific).Value = r["TYPE"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_21").Cells.Item(i + 1).Specific).Value = r["UoM"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_22").Cells.Item(i + 1).Specific).Value = r["UPrice"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_23").Cells.Item(i + 1).Specific).Value = r["Quantity"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_24").Cells.Item(i + 1).Specific).Value = r["Total"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_2").Cells.Item(i + 1).Specific).Value = r["GoiThauKey"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_3").Cells.Item(i + 1).Specific).Value = r["GoiThauName"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_4").Cells.Item(i + 1).Specific).Value = r["SubProjectKey"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_5").Cells.Item(i + 1).Specific).Value = r["SubProjectName"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_6").Cells.Item(i + 1).Specific).Value = r["StagesKey"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_7").Cells.Item(i + 1).Specific).Value = r["OpenIssuesKey"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_8").Cells.Item(i + 1).Specific).Value = r["Remarks"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_9").Cells.Item(i + 1).Specific).Value = "";
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_10").Cells.Item(i + 1).Specific).Value = r["UoM"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_11").Cells.Item(i + 1).Specific).Value = r["UPrice"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_12").Cells.Item(i + 1).Specific).Value = r["Quantity"].ToString();
                        //((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_13").Cells.Item(i + 1).Specific).Value = r["Total"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_25").Cells.Item(i + 1).Specific).Value = BType == "3" ? "100" : r["Last_Complete_Rate"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_26").Cells.Item(i + 1).Specific).Value = BType == "3" ? r["Total"].ToString() : r["Last_Complete_Amount"].ToString();
                        if (i < B.Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_1_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }
                oForm.Items.Item("20_U_E").Click();
                //Invisible Column 
                oMtx.Columns.Item("C_1_2").Visible = false;
                oMtx.Columns.Item("C_1_4").Visible = false;
                oMtx.Columns.Item("C_1_6").Visible = false;
                oMtx.Columns.Item("C_1_7").Visible = false;
                //Disable All Column except 2 last column
                oMtx.AutoResizeColumns();
                foreach (SAPbouiCOM.Column c in oMtx.Columns)
                {
                    if (c.Title != "Phần trăm hoàn thành" && c.Title != "Giá trị thực hiện")
                        c.Editable = false;
                    else
                        c.Editable = true;
                }
                oForm.Freeze(false);

                //Fill Value C - Khau tru vat tu may moc
                DataTable C = new DataTable();// Get_Data_KLTT(fProject, "C", To_Date, BPCode, BGroup, PurchaseType);
                if (NT)
                    C = Get_Data_KLTT_NT(fProject, "C", Period, BPCode, BGroup, PurchaseType);
                else
                    C = Get_Data_KLTT(fProject, "C", To_Date, BPCode, BGroup, PurchaseType);
                oForm.PaneLevel = 3;
                oForm.Freeze(true);
                oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("2_U_G").Specific;
                if (C.Rows.Count > 0)
                {
                    for (int i = 0; i < C.Rows.Count; i++)
                    {
                        DataRow r = C.Rows[i];
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_2_2").Cells.Item(i + 1).Specific).Value = r["GoiThauKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_2_3").Cells.Item(i + 1).Specific).Value = r["GoiThauName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_2_4").Cells.Item(i + 1).Specific).Value = r["GIKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_2_5").Cells.Item(i + 1).Specific).Value = r["GIRowKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_2_6").Cells.Item(i + 1).Specific).Value = r["DetailsName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_2_7").Cells.Item(i + 1).Specific).Value = r["UoM"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_2_8").Cells.Item(i + 1).Specific).Value = r["UPrice"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_2_9").Cells.Item(i + 1).Specific).Value = r["Quantity"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_2_10").Cells.Item(i + 1).Specific).Value = r["Total"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_2_11").Cells.Item(i + 1).Specific).Value = BType == "3" ? "100" : r["Last_Complete_Rate"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_2_12").Cells.Item(i + 1).Specific).Value = BType == "3" ? r["Total"].ToString() : r["Last_Complete_Amount"].ToString();
                        if (i < C.Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_2_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }
                oForm.Items.Item("20_U_E").Click();
                //Invisible Column 
                oMtx.Columns.Item("C_2_2").Visible = false;
                oMtx.Columns.Item("C_2_4").Visible = false;
                oMtx.Columns.Item("C_2_5").Visible = false;
                //Disable All Column except 2 last column
                oMtx.AutoResizeColumns();
                foreach (SAPbouiCOM.Column c in oMtx.Columns)
                {
                    if (c.Title != "Phần trăm hoàn thành" && c.Title != "Giá trị thực hiện")
                        c.Editable = false;
                    else
                        c.Editable = true;
                }
                oForm.Freeze(false);

                //Fill Value D - Khau tru bao ho lao dong
                DataTable D = new DataTable();// Get_Data_KLTT(fProject, "D", To_Date, BPCode, BGroup, PurchaseType);
                if (NT)
                    D = Get_Data_KLTT_NT(fProject, "D", Period, BPCode, BGroup, PurchaseType);
                else
                    D = Get_Data_KLTT(fProject, "D", To_Date, BPCode, BGroup, PurchaseType);
                oForm.PaneLevel = 4;
                oForm.Freeze(true);
                oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("3_U_G").Specific;
                if (D.Rows.Count > 0)
                {
                    for (int i = 0; i < D.Rows.Count; i++)
                    {
                        DataRow r = D.Rows[i];
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_3_2").Cells.Item(i + 1).Specific).Value = r["GoiThauKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_3_3").Cells.Item(i + 1).Specific).Value = r["GoiThauName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_3_4").Cells.Item(i + 1).Specific).Value = r["GIKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_3_5").Cells.Item(i + 1).Specific).Value = r["GIRowKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_3_6").Cells.Item(i + 1).Specific).Value = r["DetailsName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_3_7").Cells.Item(i + 1).Specific).Value = r["UoM"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_3_8").Cells.Item(i + 1).Specific).Value = r["UPrice"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_3_9").Cells.Item(i + 1).Specific).Value = r["Quantity"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_3_10").Cells.Item(i + 1).Specific).Value = r["Total"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_3_11").Cells.Item(i + 1).Specific).Value = BType == "3" ? "100" : r["Last_Complete_Rate"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_3_12").Cells.Item(i + 1).Specific).Value = BType == "3" ? r["Total"].ToString() : r["Last_Complete_Amount"].ToString();
                        if (i < D.Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_3_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }
                oForm.Items.Item("20_U_E").Click();
                //Invisible Column 
                oMtx.Columns.Item("C_3_2").Visible = false;
                oMtx.Columns.Item("C_3_4").Visible = false;
                oMtx.Columns.Item("C_3_5").Visible = false;
                //Disable All Column except 2 last column
                oMtx.AutoResizeColumns();
                foreach (SAPbouiCOM.Column c in oMtx.Columns)
                {
                    if (c.Title != "Phần trăm hoàn thành" && c.Title != "Giá trị thực hiện")
                        c.Editable = false;
                    else
                        c.Editable = true;
                }
                oForm.Freeze(false);

                //Fill Value E - Ho tro thi cong theo QC
                DataTable E = new DataTable();// Get_Data_KLTT(fProject, "E", To_Date, BPCode, BGroup, PurchaseType);
                if (NT)
                    E = Get_Data_KLTT_NT(fProject, "E", Period, BPCode, BGroup, PurchaseType);
                else
                    E = Get_Data_KLTT(fProject, "E", To_Date, BPCode, BGroup, PurchaseType);
                oForm.PaneLevel = 5;
                oForm.Freeze(true);
                oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("4_U_G").Specific;
                if (E.Rows.Count > 0)
                {
                    for (int i = 0; i < E.Rows.Count; i++)
                    {
                        DataRow r = E.Rows[i];
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_2").Cells.Item(i + 1).Specific).Value = r["GoiThauKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_3").Cells.Item(i + 1).Specific).Value = r["GoiThauName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_4").Cells.Item(i + 1).Specific).Value = r["SubProjectKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_5").Cells.Item(i + 1).Specific).Value = r["StagesKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_6").Cells.Item(i + 1).Specific).Value = r["OpenIssuesKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_7").Cells.Item(i + 1).Specific).Value = r["Remarks"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_8").Cells.Item(i + 1).Specific).Value = r["UoM"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_9").Cells.Item(i + 1).Specific).Value = r["UPrice"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_10").Cells.Item(i + 1).Specific).Value = r["Quantity"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_11").Cells.Item(i + 1).Specific).Value = r["Total"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_12").Cells.Item(i + 1).Specific).Value = BType == "3" ? "100" : r["Last_Complete_Rate"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_13").Cells.Item(i + 1).Specific).Value = BType == "3" ? r["Total"].ToString() : r["Last_Complete_Amount"].ToString();
                        if (i < E.Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_4_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }
                oForm.Items.Item("20_U_E").Click();
                //Invisible Column 
                oMtx.Columns.Item("C_4_2").Visible = false;
                oMtx.Columns.Item("C_4_4").Visible = false;
                oMtx.Columns.Item("C_4_5").Visible = false;
                oMtx.Columns.Item("C_4_6").Visible = false;
                //Disable All Column except 2 last column
                oMtx.AutoResizeColumns();
                foreach (SAPbouiCOM.Column c in oMtx.Columns)
                {
                    if (c.Title != "Phần trăm hoàn thành" && c.Title != "Giá trị thực hiện")
                        c.Editable = false;
                    else
                        c.Editable = true;
                }
                oForm.Freeze(false);

                //Fill Value F - Ho tro thi cong ngoai QC
                DataTable F = new DataTable();// Get_Data_KLTT(fProject, "F", To_Date, BPCode, BGroup, PurchaseType);
                if (NT)
                    F = Get_Data_KLTT_NT(fProject, "F", Period, BPCode, BGroup, PurchaseType);
                else
                    F = Get_Data_KLTT(fProject, "F", To_Date, BPCode, BGroup, PurchaseType);
                oForm.PaneLevel = 6;
                oForm.Freeze(true);
                oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("5_U_G").Specific;
                if (F.Rows.Count > 0)
                {
                    for (int i = 0; i < F.Rows.Count; i++)
                    {
                        DataRow r = F.Rows[i];
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_2").Cells.Item(i + 1).Specific).Value = r["GoiThauKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_3").Cells.Item(i + 1).Specific).Value = r["GoiThauName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_4").Cells.Item(i + 1).Specific).Value = r["SubProjectKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_5").Cells.Item(i + 1).Specific).Value = r["StagesKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_6").Cells.Item(i + 1).Specific).Value = r["OpenIssuesKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_7").Cells.Item(i + 1).Specific).Value = r["Remarks"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_8").Cells.Item(i + 1).Specific).Value = r["UoM"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_9").Cells.Item(i + 1).Specific).Value = r["UPrice"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_10").Cells.Item(i + 1).Specific).Value = r["Quantity"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_11").Cells.Item(i + 1).Specific).Value = r["Total"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_12").Cells.Item(i + 1).Specific).Value = BType == "3" ? "100" : r["Last_Complete_Rate"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_13").Cells.Item(i + 1).Specific).Value = BType == "3" ? r["Total"].ToString() : r["Last_Complete_Amount"].ToString();
                        if (i < F.Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_5_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }
                oForm.Items.Item("20_U_E").Click();
                //Invisible Column 
                oMtx.Columns.Item("C_5_2").Visible = false;
                oMtx.Columns.Item("C_5_4").Visible = false;
                oMtx.Columns.Item("C_5_5").Visible = false;
                oMtx.Columns.Item("C_5_6").Visible = false;
                //Disable All Column except 2 last column
                oMtx.AutoResizeColumns();
                foreach (SAPbouiCOM.Column c in oMtx.Columns)
                {
                    if (c.Title != "Phần trăm hoàn thành" && c.Title != "Giá trị thực hiện")
                        c.Editable = false;
                    else
                        c.Editable = true;
                }
                oForm.Freeze(false);

                //Fill Value G - Thuong phat
                DataTable G = new DataTable();// Get_Data_KLTT(fProject, "G", To_Date, BPCode, BGroup, PurchaseType);
                if (NT)
                    G = Get_Data_KLTT_NT(fProject, "G", Period, BPCode, BGroup, PurchaseType);
                else
                    G = Get_Data_KLTT(fProject, "G", To_Date, BPCode, BGroup, PurchaseType);
                oForm.PaneLevel = 7;
                oForm.Freeze(true);
                oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("6_U_G").Specific;
                if (G.Rows.Count > 0)
                {
                    for (int i = 0; i < G.Rows.Count; i++)
                    {
                        DataRow r = G.Rows[i];
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_2").Cells.Item(i + 1).Specific).Value = r["GoiThauKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_3").Cells.Item(i + 1).Specific).Value = r["GoiThauName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_4").Cells.Item(i + 1).Specific).Value = r["SubProjectKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_5").Cells.Item(i + 1).Specific).Value = r["StagesKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_6").Cells.Item(i + 1).Specific).Value = r["OpenIssuesKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_7").Cells.Item(i + 1).Specific).Value = r["Remarks"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_8").Cells.Item(i + 1).Specific).Value = r["UoM"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_9").Cells.Item(i + 1).Specific).Value = r["UPrice"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_10").Cells.Item(i + 1).Specific).Value = r["Quantity"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_11").Cells.Item(i + 1).Specific).Value = r["Total"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_12").Cells.Item(i + 1).Specific).Value = BType == "3" ? "100" : r["Last_Complete_Rate"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_13").Cells.Item(i + 1).Specific).Value = BType == "3" ? r["Total"].ToString() : r["Last_Complete_Amount"].ToString();
                        if (i < G.Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_6_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }
                oForm.Items.Item("20_U_E").Click();
                //Invisible Column 
                oMtx.Columns.Item("C_6_2").Visible = false;
                oMtx.Columns.Item("C_6_4").Visible = false;
                oMtx.Columns.Item("C_6_5").Visible = false;
                oMtx.Columns.Item("C_6_6").Visible = false;
                //Disable All Column except 2 last column
                oMtx.AutoResizeColumns();
                foreach (SAPbouiCOM.Column c in oMtx.Columns)
                {
                    if (c.Title != "Phần trăm hoàn thành" && c.Title != "Giá trị thực hiện")
                        c.Editable = false;
                    else
                        c.Editable = true;

                }
                oForm.Freeze(false);

                //Fill Value H - Phat sinh
                DataTable H = Get_Data_KLTT(fProject, "H", To_Date, BPCode, BGroup, PurchaseType);
                oForm.PaneLevel = 8;
                oForm.Freeze(true);
                oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("7_U_G").Specific;
                if (H.Rows.Count > 0)
                {
                    for (int i = 0; i < H.Rows.Count; i++)
                    {
                        DataRow r = H.Rows[i];
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_7_2").Cells.Item(i + 1).Specific).Value = r["AbsId"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_7_3").Cells.Item(i + 1).Specific).Value = r["Number"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_7_4").Cells.Item(i + 1).Specific).Value = r["AgrLineNum"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_7_5").Cells.Item(i + 1).Specific).Value = r["Status"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_7_6").Cells.Item(i + 1).Specific).Value = DateTime.Parse(r["StartDate"].ToString()).ToString("yyyyMMdd");
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_7_7").Cells.Item(i + 1).Specific).Value = r["ItemCode"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_7_8").Cells.Item(i + 1).Specific).Value = r["ItemName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_7_9").Cells.Item(i + 1).Specific).Value = r["PlanQty"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_7_10").Cells.Item(i + 1).Specific).Value = r["InvntryUom"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_7_11").Cells.Item(i + 1).Specific).Value = r["UnitPrice"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_7_12").Cells.Item(i + 1).Specific).Value = r["Total"].ToString();
                        if (i < H.Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_7_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }

                oForm.Items.Item("20_U_E").Click();
                //Invisible Column 
                oMtx.Columns.Item("C_7_2").Visible = false;
                oMtx.Columns.Item("C_7_4").Visible = false;
                oMtx.Columns.Item("C_7_6").Visible = false;
                //Disable All Column except 2 last column
                oMtx.AutoResizeColumns();
                foreach (SAPbouiCOM.Column c in oMtx.Columns)
                {
                    if (c.Title != "Phần trăm hoàn thành" && c.Title != "Giá trị thực hiện")
                        c.Editable = false;
                    else
                        c.Editable = true;

                }
                oForm.Freeze(false);

                //Fill Value I - Phe duyet
                oForm.PaneLevel = 9;
                oForm.Freeze(true);
                DataTable APPRV= null;
                if (string.IsNullOrEmpty(QLNT) || QLNT=="")
                    APPRV = GetList_Approve(BGroup);
                else
                    APPRV = GetList_Approve(BGroup,"Y");

                if (APPRV.Rows.Count > 0)
                {
                    oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("8_U_G").Specific;
                    for (int i = 0; i < APPRV.Rows.Count; i++)
                    {
                        DataRow r = APPRV.Rows[i];
                        ((SAPbouiCOM.ComboBox)oMtx.Columns.Item("C_8_2").Cells.Item(i + 1).Specific).Select(r["LEVEL"].ToString(),SAPbouiCOM.BoSearchKey.psk_ByValue);
                        if (r["Position"].ToString() != "" && !string.IsNullOrEmpty(r["Position"].ToString()))
                            ((SAPbouiCOM.ComboBox)oMtx.Columns.Item("C_8_5").Cells.Item(i + 1).Specific).Select(r["Position"].ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                        if (i < APPRV.Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_8_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }
                oForm.Freeze(false);

                //Fill Value K - Phat sinh NEW Vàng
                DataTable K = new DataTable();// Get_Data_KLTT(fProject, "K", To_Date, BPCode, BGroup, PurchaseType);
                if (NT)
                    K = Get_Data_KLTT_NT(fProject, "K", Period, BPCode, BGroup, PurchaseType);
                else
                    K = Get_Data_KLTT(fProject, "K", To_Date, BPCode, BGroup, PurchaseType);
                oForm.PaneLevel = 10;
                oForm.Freeze(true);
                oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item("9_U_G").Specific;
                if (K.Rows.Count > 0)
                {
                    for (int i = 0; i < K.Rows.Count; i++)
                    {
                        DataRow r = K.Rows[i];
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_2").Cells.Item(i + 1).Specific).Value = r["GoiThauKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_3").Cells.Item(i + 1).Specific).Value = r["GoiThauName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_4").Cells.Item(i + 1).Specific).Value = r["GRPOKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_5").Cells.Item(i + 1).Specific).Value = r["GRPORowKey"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_6").Cells.Item(i + 1).Specific).Value = r["DetailsName"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_7").Cells.Item(i + 1).Specific).Value = r["TYPE"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_8").Cells.Item(i + 1).Specific).Value = r["DetailsWork"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_9").Cells.Item(i + 1).Specific).Value = r["U_ParentID1"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_10").Cells.Item(i + 1).Specific).Value = r["U_ParentID2"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_11").Cells.Item(i + 1).Specific).Value = r["U_ParentID3"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_12").Cells.Item(i + 1).Specific).Value = r["U_ParentID4"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_13").Cells.Item(i + 1).Specific).Value = r["U_ParentID5"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_14").Cells.Item(i + 1).Specific).Value = r["Name1"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_15").Cells.Item(i + 1).Specific).Value = r["Name2"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_16").Cells.Item(i + 1).Specific).Value = r["Name3"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_17").Cells.Item(i + 1).Specific).Value = r["Name4"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_18").Cells.Item(i + 1).Specific).Value = r["Name5"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_19").Cells.Item(i + 1).Specific).Value = r["UoM"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_20").Cells.Item(i + 1).Specific).Value = r["UPrice"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_21").Cells.Item(i + 1).Specific).Value = r["Quantity"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_22").Cells.Item(i + 1).Specific).Value = r["Total"].ToString();

                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_23").Cells.Item(i + 1).Specific).Value = BType == "3" ? "100" : r["Last_Complete_Rate"].ToString();
                        ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_24").Cells.Item(i + 1).Specific).Value = BType == "3" ? r["Total"].ToString() : r["Last_Complete_Amount"].ToString();
                        if (i < K.Rows.Count - 1)
                        {
                            oMtx.AddRow();
                            ((SAPbouiCOM.EditText)oMtx.Columns.Item("C_9_1").Cells.Item(i + 2).Specific).Value = (i + 2).ToString();
                        }
                    }
                }
                oForm.Items.Item("20_U_E").Click();
                //Invisible Column 
                //oMtx.Columns.Item("C_1_2").Visible = false;
                //oMtx.Columns.Item("C_1_4").Visible = false;
                //oMtx.Columns.Item("C_1_6").Visible = false;
                //oMtx.Columns.Item("C_1_7").Visible = false;
                //Disable All Column except 2 last column
                oMtx.AutoResizeColumns();
                foreach (SAPbouiCOM.Column c in oMtx.Columns)
                {
                    if (c.Title != "Phần trăm hoàn thành" && c.Title != "Giá trị thực hiện")
                        c.Editable = false;
                    else
                        c.Editable = true;
                }
                oForm.Freeze(false);

                CheckBox0.Checked = false;
                oApp.StatusBar.SetText("Load Data Completed", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        //View
        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (Grid0.Rows.SelectedRows.Count == 1)
                {
                    oApp.ActivateMenuItem("KLTT");
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
                    frm.EnableMenu("KLTT_Remove_Line", false);
                    frm.EnableMenu("KLTT_Add_Line", false);
                    frm.Items.Item("0_U_G").Enabled = false;
                    frm.Items.Item("1_U_G").Enabled = false;
                    frm.Items.Item("2_U_G").Enabled = false;
                    frm.Items.Item("3_U_G").Enabled = false;
                    frm.Items.Item("4_U_G").Enabled = false;
                    frm.Items.Item("5_U_G").Enabled = false;
                    frm.Items.Item("6_U_G").Enabled = false;
                    //Disable Header
                    frm.Items.Item("20_U_E").Enabled = false;
                    frm.Items.Item("21_U_E").Enabled = false;
                    frm.Items.Item("22_U_E").Enabled = false;
                    frm.Items.Item("23_U_E").Enabled = false;
                    frm.Items.Item("24_U_E").Enabled = false;
                    frm.Items.Item("25_U_E").Enabled = false;
                    frm.Items.Item("26_U_E").Enabled = false;
                    frm.Items.Item("27_U_E").Enabled = false;
                    frm.Items.Item("28_U_E").Enabled = false;
                    frm.Items.Item("29_U_Cb").Enabled = false;
                    frm.Items.Item("30_U_Cb").Enabled = false;
                    frm.Items.Item("32_U_Cb").Enabled = false;
                    frm.Items.Item("31_U_E").Enabled = false;
                    //Disable all column Matrix
                    Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("0_U_G").Specific, false);
                    Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("1_U_G").Specific, false);
                    Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("2_U_G").Specific, false);
                    Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("3_U_G").Specific, false);
                    Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("4_U_G").Specific, false);
                    Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("5_U_G").Specific, false);
                    Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("6_U_G").Specific, false);
                    Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("7_U_G").Specific, false);
                    Editable_Column_Matrix((SAPbouiCOM.Matrix)frm.Items.Item("9_U_G").Specific, false);

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

        //Edit
        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (Grid0.Rows.SelectedRows.Count == 1)
                {
                    oApp.ActivateMenuItem("KLTT");
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
                    frm.EnableMenu("1304", false);  // Close
                    frm.EnableMenu("KLTT_Remove_Line", false);
                    frm.EnableMenu("KLTT_Add_Line", false);
                    //frm.Items.Item("0_U_G").Enabled = false;
                    //frm.Items.Item("1_U_G").Enabled = false;
                    //frm.Items.Item("2_U_G").Enabled = false;
                    //frm.Items.Item("3_U_G").Enabled = false;
                    //frm.Items.Item("4_U_G").Enabled = false;
                    //frm.Items.Item("5_U_G").Enabled = false;
                    //frm.Items.Item("6_U_G").Enabled = false;
                    //Disable Header
                    frm.Items.Item("20_U_E").Enabled = false;
                    //frm.Items.Item("21_U_E").Enabled = false;
                    //frm.Items.Item("22_U_E").Enabled = false;
                    //frm.Items.Item("23_U_E").Enabled = false;
                    frm.Items.Item("24_U_E").Enabled = false;
                    //frm.Items.Item("25_U_E").Enabled = false;
                    frm.Items.Item("26_U_E").Enabled = false;
                    frm.Items.Item("27_U_E").Enabled = false;
                    //Disable all column Matrix
                    Editable_Column_Matrix_EditMode((SAPbouiCOM.Matrix)frm.Items.Item("0_U_G").Specific);
                    Editable_Column_Matrix_EditMode((SAPbouiCOM.Matrix)frm.Items.Item("1_U_G").Specific);
                    Editable_Column_Matrix_EditMode((SAPbouiCOM.Matrix)frm.Items.Item("2_U_G").Specific);
                    Editable_Column_Matrix_EditMode((SAPbouiCOM.Matrix)frm.Items.Item("3_U_G").Specific);
                    Editable_Column_Matrix_EditMode((SAPbouiCOM.Matrix)frm.Items.Item("4_U_G").Specific);
                    Editable_Column_Matrix_EditMode((SAPbouiCOM.Matrix)frm.Items.Item("5_U_G").Specific);
                    Editable_Column_Matrix_EditMode((SAPbouiCOM.Matrix)frm.Items.Item("6_U_G").Specific);
                    Editable_Column_Matrix_EditMode((SAPbouiCOM.Matrix)frm.Items.Item("7_U_G").Specific);
                    Editable_Column_Matrix_EditMode((SAPbouiCOM.Matrix)frm.Items.Item("9_U_G").Specific);
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

        //Load Data
        private void Button3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (!String.IsNullOrEmpty(EditText1.Value))
            {
                if (!string.IsNullOrEmpty(ComboBox0.Value))
                {
                    Load_Grid_Period(this.EditText1.Value.Trim(), ComboBox0.Selected.Value, ComboBox1.Selected.Value, ComboBox3.Selected.Value);
                }
            }
        }

        //View Bill
        private void Button4_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            object misvalue = System.Reflection.Missing.Value;
            try
            {
                if (Grid0.Rows.SelectedRows.Count == 1)
                {
                    //Get Data
                    int DocEntry = 0;
                    string docnum = Grid0.DataTable.GetValue("Document Number", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                    int.TryParse(docnum, out DocEntry);
                    string projectKey = Grid0.DataTable.GetValue("Financial Project", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();

                    //Get Header Info
                    DataTable dt_tmp = new DataTable();
                    //Select a.*,b.PrjName from [@KLTT] a left join OPRJ b on a.U_ProjectNo = b.PrjCode
                    string sql_cmd = string.Format("Select a.*,b.PrjName,(Select GroupCode from OCRD where CardCode=a.U_BpCode ) as 'BPGCode' from [@KLTT] a left join OPRJ b on a.U_FIPROJECT = b.PrjCode where a.U_FIPROJECT='{0}' and a.DocEntry = {1};", projectKey, docnum);
                    try
                    {
                        SqlCommand cmd = new SqlCommand(sql_cmd, conn);
                        conn.Open();
                        SqlDataReader dr = cmd.ExecuteReader();
                        dt_tmp.Load(dr);
                    }
                    catch
                    {

                    }
                    finally
                    {
                        conn.Close();
                    }
                    if (dt_tmp.Rows.Count == 1)
                    {
                        //Header Report
                        string financialproject = dt_tmp.Rows[0]["U_FIPROJECT"].ToString();
                        //string projectNo = dt_tmp.Rows[0]["U_ProjectNo"].ToString();
                        string projectName = dt_tmp.Rows[0]["PrjName"].ToString();
                        string bp = dt_tmp.Rows[0]["U_BPCode"].ToString();
                        string bpname = dt_tmp.Rows[0]["U_BPName"].ToString();
                        int period = 1;
                        int.TryParse(dt_tmp.Rows[0]["U_Period"].ToString(), out period);
                        DateTime frdate = DateTime.Today;
                        DateTime.TryParse(dt_tmp.Rows[0]["U_DateFrom"].ToString(), out frdate);
                        DateTime todate = DateTime.Today;
                        DateTime.TryParse(dt_tmp.Rows[0]["U_DateTo"].ToString(), out todate);
                        string BGroup = dt_tmp.Rows[0]["U_BGroup"].ToString();
                        string BType = dt_tmp.Rows[0]["U_BType"].ToString();
                        string PuType = dt_tmp.Rows[0]["U_PUTYPE"].ToString();
                        string QLNT = dt_tmp.Rows[0]["U_BPCode2"].ToString();
                        string BPGCode = dt_tmp.Rows[0]["BPGCode"].ToString();
                        float PhiQL = 0;
                        float.TryParse(dt_tmp.Rows[0]["U_PTQuanly"].ToString(), out PhiQL);

                        DataTable A = Load_Data_KLTT(financialproject, "A", period, bp, BGroup, DocEntry, PuType);
                        DataTable B = Load_Data_KLTT(financialproject, "B", period, bp, BGroup, DocEntry, PuType);
                        DataTable C = Load_Data_KLTT(financialproject, "C", period, bp, BGroup, DocEntry, PuType);
                        DataTable D = Load_Data_KLTT(financialproject, "D", period, bp, BGroup, DocEntry, PuType);
                        DataTable E = Load_Data_KLTT(financialproject, "E", period, bp, BGroup, DocEntry, PuType);
                        DataTable F = Load_Data_KLTT(financialproject, "F", period, bp, BGroup, DocEntry, PuType);
                        DataTable G = Load_Data_KLTT(financialproject, "G", period, bp, BGroup, DocEntry, PuType);
                        DataTable H = Load_Data_KLTT(financialproject, "H", period, bp, BGroup, DocEntry, PuType);
                        DataTable K = Load_Data_KLTT(financialproject, "K", period, bp, BGroup, DocEntry, PuType);
                        DataTable Additional_Info = Load_Data_KLTT_AI(financialproject, period, bp, todate, BGroup, PuType);
                        DataTable Total_PP = Load_Data_KLTT(financialproject, "SPP", period - 1, bp, BGroup, DocEntry, PuType);

                        //Get List Goi Thau
                        DataTable tb_goithau = A.AsDataView().ToTable(true, new string[] { "U_Sub1", "U_Sub1Name" });

                        //Start Excel and get Application object.
                        oXL = new Microsoft.Office.Interop.Excel.Application();
                        oXL.Visible = true;
                        //Open Template
                        if (QLNT != "" && !string.IsNullOrEmpty(QLNT))
                            oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\Template_KLTT_NT.xlsx");
                        else
                            oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\Template_KLTT.xlsx");
                        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                        //Fill Header
                        //Tilte
                        if(BType =="3")
                            oSheet.Cells[4, 3] = "BẢNG KHỐI LƯỢNG QUYẾT TOÁN";
                        //Project Name
                        oSheet.Cells[1, 4] = projectName;
                        //System Date
                        oSheet.Cells[2, 4] = DateTime.Today.ToString("dd/MM/yyyy");
                        //BP Name
                        oSheet.Cells[3, 4] = bp + " - " + bpname;
                        //From Date - To Date
                        oSheet.Cells[5, 3] = string.Format("Từ ngày {0} đến ngày {1}", frdate.ToString("dd/MM/yyyy"), todate.ToString("dd/MM/yyyy"));
                        //Period
                        oSheet.Cells[5, 9] = "Kỳ: " + period.ToString();

                        //Write Details to Excel
                        int current_rownum = 8;
                        string subprojectkey = "";
                        int Group_No = 0;
                        int Detail_No = 1;
                        List<int> Group_No_RowNum = new List<int>();
                        List<int> Group_No_RowNum2 = new List<int>();
                        List<int> Group_No_RowNum3 = new List<int>();
                        List<int> Group_No_RowNum4 = new List<int>();
                        List<int> Group_No_RowNum5 = new List<int>();
                        List<int> Section_RowNum = new List<int>();

                        //A - Khoi luong cong viec theo hop dong
                        #region A
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Cells[current_rownum, 1] = "A";
                        oSheet.Cells[current_rownum, 2] = "KHỐI LƯỢNG THEO HỢP ĐỒNG NCC/NTP/ĐTC";
                        Section_RowNum.Add(current_rownum);
                        current_rownum++;
                        
                        #region LV1

                        if (tb_goithau.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_goithau.Rows[0]["U_Sub1"].ToString()))
                        {
                            foreach (DataRow r in tb_goithau.Rows)
                            {
                                Group_No_RowNum2.Clear();
                                Group_No_RowNum3.Clear();
                                Group_No_RowNum4.Clear();
                                Group_No_RowNum5.Clear();
                                //Goi Thau
                                List<int> Gr_element = new List<int>();
                                oSheet.Cells[current_rownum, 1] = Group_No_RowNum.Count + 1;
                                oSheet.Cells[current_rownum, 2] = r["U_Sub1Name"].ToString();
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 8]).Formula = string.Format(@"=I{0}/G{1}", current_rownum, current_rownum);
                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                Group_No_RowNum.Add(current_rownum);
                                current_rownum++;

                                //Gang chung tu o Sub1 - Goi thau
                                foreach (DataRow t in A.Select("U_Sub1 ='" + r["U_Sub1"] + "' and ISNULL(U_Sub2,'') =''"))
                                {
                                    oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                    oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                    oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                    oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                    oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                    oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                    oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                    oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                    Gr_element.Add(current_rownum);
                                    current_rownum++;
                                }

                                //Group level 2
                                #region LV2
                                DataTable tb_lv2 = Get_List_Sub_Level(int.Parse(docnum), int.Parse(r["U_Sub1"].ToString()), 2, "A");
                                if (tb_lv2.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_lv2.Rows[0]["U_Sub2"].ToString()))
                                {
                                    foreach (DataRow r_lv2 in tb_lv2.Rows)
                                    {
                                        List<int> Gr2_element = new List<int>();
                                        if (!String.IsNullOrEmpty(r_lv2["U_Sub2"].ToString()))
                                        {
                                            oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(255, 217, 102);
                                            oSheet.Cells[current_rownum, 1] = string.Format("{0}.{1}", Group_No_RowNum.Count, Group_No_RowNum2.Count + 1);// (Group_No_RowNum.Count + 1);
                                            oSheet.Cells[current_rownum, 2] = r_lv2["U_Sub2Name"].ToString();
                                            Group_No_RowNum2.Add(current_rownum);
                                            Gr_element.Add(current_rownum);
                                            current_rownum++;

                                            foreach (DataRow t in A.Select("U_Sub2 ='" + r_lv2["U_Sub2"] + "' and ISNULL(U_Sub3,'') =''"))
                                            {
                                                oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                                oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                                oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                                oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                                oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                                oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                                oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                                oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                                Gr2_element.Add(current_rownum);
                                                current_rownum++;
                                            }

                                            //Group level 3
                                            #region LV3
                                            DataTable tb_lv3 = Get_List_Sub_Level(int.Parse(docnum), int.Parse(r_lv2["U_Sub2"].ToString()), 3, "A");
                                            if (tb_lv3.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_lv3.Rows[0]["U_Sub3"].ToString()))
                                            {
                                                foreach (DataRow r_lv3 in tb_lv3.Rows)
                                                {
                                                    List<int> Gr3_element = new List<int>();
                                                    if (!String.IsNullOrEmpty(r_lv3["U_Sub3"].ToString()))
                                                    {
                                                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(255, 242, 204);
                                                        oSheet.Cells[current_rownum, 2] = r_lv3["U_Sub3Name"].ToString();
                                                        oSheet.Cells[current_rownum, 1] = string.Format("{0}.{1}.{2}", Group_No_RowNum.Count, Group_No_RowNum2.Count, Group_No_RowNum3.Count + 1);
                                                        Gr2_element.Add(current_rownum);
                                                        Group_No_RowNum3.Add(current_rownum);
                                                        current_rownum++;

                                                        foreach (DataRow t in A.Select("U_Sub3 ='" + r_lv3["U_Sub3"] + "' and ISNULL(U_Sub4,'') =''"))
                                                        {
                                                            oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                                            oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                                            oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                                            oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                                            oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                                            oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                                            oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                                            oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                                            Gr3_element.Add(current_rownum);
                                                            current_rownum++;
                                                        }

                                                        //Group level 4
                                                        #region LV4
                                                        DataTable tb_lv4 = Get_List_Sub_Level(int.Parse(docnum), int.Parse(r_lv3["U_Sub3"].ToString()), 4, "A");
                                                        if (tb_lv4.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_lv4.Rows[0]["U_Sub4"].ToString()))
                                                        {

                                                            foreach (DataRow r_lv4 in tb_lv4.Rows)
                                                            {
                                                                List<int> Gr4_element = new List<int>();
                                                                if (!String.IsNullOrEmpty(r_lv4["U_Sub4"].ToString()))
                                                                {
                                                                    oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);
                                                                    oSheet.Cells[current_rownum, 2] = r_lv4["U_Sub4Name"].ToString();
                                                                    Group_No_RowNum4.Add(current_rownum);
                                                                    Gr3_element.Add(current_rownum);
                                                                    current_rownum++;

                                                                    foreach (DataRow t in A.Select("U_Sub4 ='" + r_lv4["U_Sub4"] + "' and ISNULL(U_Sub5,'') =''"))
                                                                    {
                                                                        oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                                                        oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                                                        oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                                                        oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                                                        oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                                                        oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                                                        oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                                                        oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                                                        Gr4_element.Add(current_rownum);
                                                                        current_rownum++;

                                                                    }
                                                                    //Group level 5
                                                                    #region LV5
                                                                    DataTable tb_lv5 = Get_List_Sub_Level(int.Parse(docnum), int.Parse(r_lv4["U_Sub4"].ToString()), 5, "A");
                                                                    if (tb_lv5.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_lv5.Rows[0]["U_Sub5"].ToString()))
                                                                    {
                                                                        foreach (DataRow r_lv5 in tb_lv5.Rows)
                                                                        {
                                                                            if (!String.IsNullOrEmpty(r_lv5["U_Sub5"].ToString()))
                                                                            {
                                                                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(214, 220, 218);
                                                                                oSheet.Cells[current_rownum, 2] = r_lv5["U_Sub5Name"].ToString();
                                                                                Gr4_element.Add(current_rownum);
                                                                                Group_No_RowNum5.Add(current_rownum);
                                                                                current_rownum++;

                                                                                //Detail
                                                                                foreach (DataRow t in A.Select("U_Sub5 ='" + r_lv5["U_Sub5"] + "'"))
                                                                                {
                                                                                    oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                                                                    oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                                                                    oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                                                                    oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                                                                    oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                                                                    oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                                                                    oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                                                                    oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                                                                    current_rownum++;

                                                                                }
                                                                            }

                                                                        }
                                                                        //Total Level 5
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1}", Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1, current_rownum - 1); //"=SUM(E" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":E" + (current_rownum - 1).ToString() + ")";
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1}", Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1, current_rownum - 1); //"=SUM(G" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":G" + (current_rownum - 1).ToString() + ")";
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1}", Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1, current_rownum - 1); //"=SUM(I" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":I" + (current_rownum - 1).ToString() + ")";
                                                                    }

                                                                    #endregion

                                                                    //Total Level 4
                                                                    if (Gr4_element.Count > 0)
                                                                    {
                                                                        //string cell_sum_tt = "";
                                                                        //string cell_sum_gtth = "";
                                                                        //string cell_sum_gtth2 = "";
                                                                        //int temp = 1;
                                                                        //foreach (int t in Gr4_element)
                                                                        //{
                                                                        //    if (temp < Gr4_element.Count)
                                                                        //    {
                                                                        //        cell_sum_tt += "E" + t + ",";
                                                                        //        cell_sum_gtth += "G" + t + ",";
                                                                        //        cell_sum_gtth2 += "I" + t + ",";
                                                                        //        temp++;
                                                                        //    }
                                                                        //    else
                                                                        //    {
                                                                        //        cell_sum_tt += "E" + t;
                                                                        //        cell_sum_gtth += "G" + t;
                                                                        //        cell_sum_gtth2 += "I" + t;
                                                                        //    }
                                                                        //}
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr4_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Gr4_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Gr4_element[0], current_rownum - 1);  //string.Format("=SUM({0})", cell_sum_gtth2);
                                                                    }
                                                                }

                                                            }
                                                            //if (Group_No_RowNum5.Count > 0)
                                                            //{
                                                            //    string cell_sum_tt = "";
                                                            //    string cell_sum_gtth = "";
                                                            //    int temp = 1;
                                                            //    foreach (int t in Group_No_RowNum5)
                                                            //    {
                                                            //        if (temp < Group_No_RowNum5.Count)
                                                            //        {
                                                            //            cell_sum_tt += "E" + t + ",";
                                                            //            cell_sum_gtth += "G" + t + ",";
                                                            //            temp++;
                                                            //        }

                                                            //        else
                                                            //        {
                                                            //            cell_sum_tt += "E" + t;
                                                            //            cell_sum_gtth += "G" + t;
                                                            //        }
                                                            //    }
                                                            //    //Total Level 4
                                                            //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum4.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                                            //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum4.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                                            //}

                                                        }
                                                        #endregion

                                                        //Total level 3
                                                        if (Gr3_element.Count > 0)
                                                        {
                                                            //string cell_sum_tt = "";
                                                            //string cell_sum_gtth = "";
                                                            //string cell_sum_gtth2 = "";
                                                            //int temp = 1;
                                                            //foreach (int t in Gr3_element)
                                                            //{
                                                            //    if (temp < Gr3_element.Count)
                                                            //    {
                                                            //        cell_sum_tt += "E" + t + ",";
                                                            //        cell_sum_gtth += "G" + t + ",";
                                                            //        cell_sum_gtth2 += "I" + t + ",";
                                                            //        temp++;
                                                            //    }

                                                            //    else
                                                            //    {
                                                            //        cell_sum_tt += "E" + t;
                                                            //        cell_sum_gtth += "G" + t;
                                                            //        cell_sum_gtth2 += "I" + t;
                                                            //    }
                                                            //}
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr3_element[0], current_rownum - 1); // string.Format("=SUM({0})", cell_sum_tt);
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Gr3_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Gr3_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion

                                            //Total level 2
                                            if (Gr2_element.Count > 0)
                                            {
                                                //string cell_sum_tt = "";
                                                //string cell_sum_gtth = "";
                                                //string cell_sum_gtth2 = "";
                                                //int temp = 1;
                                                //foreach (int t in Gr2_element)
                                                //{
                                                //    if (temp < Gr2_element.Count)
                                                //    {
                                                //        cell_sum_tt += "E" + t + ",";
                                                //        cell_sum_gtth += "G" + t + ",";
                                                //        cell_sum_gtth2 += "I" + t + ",";
                                                //        temp++;
                                                //    }

                                                //    else
                                                //    {
                                                //        cell_sum_tt += "E" + t;
                                                //        cell_sum_gtth += "G" + t;
                                                //        cell_sum_gtth2 += "I" + t;
                                                //    }
                                                //}
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr2_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Gr2_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Gr2_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                            }
                                            Group_No_RowNum3.Clear();
                                        }
                                    }
                                }
                                #endregion

                                //Total Goi Thau LV 1
                                if (Gr_element.Count > 0)
                                {
                                    string cell_sum_tt = "";
                                    string cell_sum_gtth = "";
                                    string cell_sum_gtth2 = "";
                                    int temp = 1;
                                    foreach (int t in Gr_element)
                                    {
                                        if (temp < Gr_element.Count)
                                        {
                                            cell_sum_tt += "E" + t + ",";
                                            cell_sum_gtth += "G" + t + ",";
                                            cell_sum_gtth2 += "I" + t + ",";
                                            temp++;
                                        }

                                        else
                                        {
                                            cell_sum_tt += "E" + t;
                                            cell_sum_gtth += "G" + t;
                                            cell_sum_gtth2 += "I" + t;
                                        }
                                    }
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Gr_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Gr_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                }

                                #region Old
                                //foreach (DataRow r_lv2 in A.Select("U_Sub1 =" + r["U_Sub1"] + "and U_Sub2 is null"))
                                //{
                                //    //Print Details
                                //    //oSheet.Cells[current_rownum, 1] = Group_No + "." + Detail_No;
                                //    oSheet.Cells[current_rownum, 2] = r_lv2["U_GPDetailsName"];
                                //    oSheet.Cells[current_rownum, 3] = r_lv2["U_CTCV"];
                                //    oSheet.Cells[current_rownum, 4] = r_lv2["U_UoM"];
                                //    oSheet.Cells[current_rownum, 5] = r_lv2["U_Quantity"];
                                //    oSheet.Cells[current_rownum, 6] = r_lv2["U_UPrice"];
                                //    oSheet.Cells[current_rownum, 7] = r_lv2["U_Sum"];
                                //    oSheet.Cells[current_rownum, 8] = r_lv2["U_CompleteRate"];
                                //    oSheet.Cells[current_rownum, 9] = r_lv2["U_CompleteAmount"];
                                //    current_rownum++;
                                //    //Detail_No++;
                                //}

                                //foreach (DataRow r_lv2 in A.Select("U_Sub1 =" + r["U_Sub1"] + " and U_Sub2 is not null"))
                                //{
                                //    oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                                //    oSheet.Cells[current_rownum, 2] = r_lv2["U_Sub2Name"].ToString();
                                //    current_rownum++;
                                //    foreach (DataRow r_lv3 in A.Select("U_Sub1 =" + r["U_Sub1"] +" and U_Sub2 = " + r_lv2["U_Sub2"] + " and U_Sub3 is null"))
                                //    {
                                //        //Print Details
                                //        //oSheet.Cells[current_rownum, 1] = Group_No + "." + Detail_No;
                                //        oSheet.Cells[current_rownum, 2] = r_lv3["U_GPDetailsName"];
                                //        oSheet.Cells[current_rownum, 3] = r_lv3["U_CTCV"];
                                //        oSheet.Cells[current_rownum, 4] = r_lv3["U_UoM"];
                                //        oSheet.Cells[current_rownum, 5] = r_lv3["U_Quantity"];
                                //        oSheet.Cells[current_rownum, 6] = r_lv3["U_UPrice"];
                                //        oSheet.Cells[current_rownum, 7] = r_lv3["U_Sum"];
                                //        oSheet.Cells[current_rownum, 8] = r_lv3["U_CompleteRate"];
                                //        oSheet.Cells[current_rownum, 9] = r_lv3["U_CompleteAmount"];
                                //        current_rownum++;
                                //        //Detail_No++;
                                //    }

                                //    //foreach (DataRow r_lv3 in A.Select("U_Sub1 =" + r["U_Sub1"] +" and U_Sub2 = " + r_lv2["U_Sub2"] + " and U_Sub3 is not null"))
                                //    //{
                                //    //    oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                //    //    oSheet.Cells[current_rownum, 2] = r_lv3["U_Sub3Name"].ToString();
                                //    //    current_rownum++;

                                //    //    foreach (DataRow r_lv4 in A.Select("U_Sub1 =" + r["U_Sub1"] +" and U_Sub2 = " + r_lv2["U_Sub2"] + " and U_Sub3 =" + r_lv3["U_Sub3"] + " and U_Sub4 is null"))
                                //    //    {
                                //    //        //Print Details
                                //    //        //oSheet.Cells[current_rownum, 1] = Group_No + "." + Detail_No;
                                //    //        oSheet.Cells[current_rownum, 2] = r_lv4["U_GPDetailsName"];
                                //    //        oSheet.Cells[current_rownum, 3] = r_lv4["U_CTCV"];
                                //    //        oSheet.Cells[current_rownum, 4] = r_lv4["U_UoM"];
                                //    //        oSheet.Cells[current_rownum, 5] = r_lv4["U_Quantity"];
                                //    //        oSheet.Cells[current_rownum, 6] = r_lv4["U_UPrice"];
                                //    //        oSheet.Cells[current_rownum, 7] = r_lv4["U_Sum"];
                                //    //        oSheet.Cells[current_rownum, 8] = r_lv4["U_CompleteRate"];
                                //    //        oSheet.Cells[current_rownum, 9] = r_lv4["U_CompleteAmount"];
                                //    //        current_rownum++;
                                //    //    }

                                //    //    foreach (DataRow r_lv4 in A.Select("U_Sub1 =" + r["U_Sub1"] +" and U_Sub2 = " + r_lv2["U_Sub2"] + " and U_Sub3 =" + r_lv3["U_Sub3"] + " and U_Sub4 is not null"))
                                //    //    {
                                //    //        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                                //    //        oSheet.Cells[current_rownum, 2] = r_lv4["U_Sub4Name"].ToString();
                                //    //        current_rownum++;

                                //    //        foreach (DataRow r_lv5 in A.Select("U_Sub1 =" + r["U_Sub1"] +" and U_Sub2 = " + r_lv2["U_Sub2"] + " and U_Sub3 =" + r_lv3["U_Sub3"] + " and U_Sub4 =" + r_lv4["U_Sub4"] + "and U_Sub5 is null"))
                                //    //        {
                                //    //            oSheet.Cells[current_rownum, 2] = r_lv5["U_GPDetailsName"];
                                //    //            oSheet.Cells[current_rownum, 3] = r_lv5["U_CTCV"];
                                //    //            oSheet.Cells[current_rownum, 4] = r_lv5["U_UoM"];
                                //    //            oSheet.Cells[current_rownum, 5] = r_lv5["U_Quantity"];
                                //    //            oSheet.Cells[current_rownum, 6] = r_lv5["U_UPrice"];
                                //    //            oSheet.Cells[current_rownum, 7] = r_lv5["U_Sum"];
                                //    //            oSheet.Cells[current_rownum, 8] = r_lv5["U_CompleteRate"];
                                //    //            oSheet.Cells[current_rownum, 9] = r_lv5["U_CompleteAmount"];
                                //    //            current_rownum++;
                                //    //        }

                                //    //        foreach (DataRow r_lv5 in A.Select("U_Sub1 =" + r["U_Sub1"] +" and U_Sub2 = " + r_lv2["U_Sub2"] + " and U_Sub3 =" + r_lv3["U_Sub3"] + " and U_Sub4 =" + r_lv4["U_Sub4"] + "and U_Sub5 is not null"))
                                //    //        {
                                //    //            oSheet.Cells[current_rownum, 2] = r_lv5["U_GPDetailsName"];
                                //    //            oSheet.Cells[current_rownum, 3] = r_lv5["U_CTCV"];
                                //    //            oSheet.Cells[current_rownum, 4] = r_lv5["U_UoM"];
                                //    //            oSheet.Cells[current_rownum, 5] = r_lv5["U_Quantity"];
                                //    //            oSheet.Cells[current_rownum, 6] = r_lv5["U_UPrice"];
                                //    //            oSheet.Cells[current_rownum, 7] = r_lv5["U_Sum"];
                                //    //            oSheet.Cells[current_rownum, 8] = r_lv5["U_CompleteRate"];
                                //    //            oSheet.Cells[current_rownum, 9] = r_lv5["U_CompleteAmount"];
                                //    //            current_rownum++;
                                //    //        }

                                //    //    }

                                //    //}
                                //}

                                //subprojectkey = "";
                                //Group_No = 0;
                                //Detail_No = 1;
                                //Group_No_RowNum.Clear();

                                //foreach (DataRow rA in A.Select("U_GoiThauKey=" + r["U_GoiThauKey"]))
                                //{
                                //    if (subprojectkey != rA["U_SubProjectKey"].ToString())
                                //    {
                                //        //Print group name
                                //        Group_No++;
                                //        Detail_No = 1;
                                //        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(198, 224, 180);
                                //        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //        oSheet.Cells[current_rownum, 1] = Group_No.ToString();
                                //        oSheet.Cells[current_rownum, 2] = rA["U_SubProjectName"];
                                //        subprojectkey = rA["U_SubProjectKey"].ToString();
                                //        //Sum thanh tien - Sum Gia tri thuc hien
                                //        int Detail_Row_Count = (int)A.Compute("Count(U_SubProjectKey)", string.Format("U_GoiThauKey={0} and U_SubProjectKey={1}", r["U_GoiThauKey"], rA["U_SubProjectKey"]));
                                //        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = "=SUM(G" + (current_rownum + 1).ToString() + ":G" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 9]).Formula = "=SUM(I" + (current_rownum + 1).ToString() + ":I" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //        Group_No_RowNum.Add(current_rownum);
                                //        current_rownum++;
                                //    }
                                //    //Print Details
                                //    oSheet.Cells[current_rownum, 1] = Group_No + "." + Detail_No;
                                //    oSheet.Cells[current_rownum, 2] = rA["U_GPDetailsName"];
                                //    oSheet.Cells[current_rownum, 3] = rA["U_CTCV"];
                                //    oSheet.Cells[current_rownum, 4] = rA["U_UoM"];
                                //    oSheet.Cells[current_rownum, 5] = rA["U_Quantity"];
                                //    oSheet.Cells[current_rownum, 6] = rA["U_UPrice"];
                                //    oSheet.Cells[current_rownum, 7] = rA["U_Sum"];
                                //    oSheet.Cells[current_rownum, 8] = rA["U_CompleteRate"];
                                //    oSheet.Cells[current_rownum, 9] = rA["U_CompleteAmount"];
                                //    current_rownum++;
                                //    Detail_No++;
                                //}
                                //if (Group_No_RowNum.Count > 0)
                                //{
                                //    string cell_sum_tt = "";
                                //    string cell_sum_gtth = "";
                                //    int temp = 1;
                                //    foreach (int t in Group_No_RowNum)
                                //    {
                                //        if (temp < Group_No_RowNum.Count)
                                //        {
                                //            cell_sum_tt += "G" + t + ",";
                                //            cell_sum_gtth += "I" + t + ",";
                                //            temp++;
                                //        }

                                //        else
                                //        {
                                //            cell_sum_tt += "G" + t;
                                //            cell_sum_gtth += "I" + t;
                                //        }
                                //    }
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 7]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                //}

                                //B - Khối lượng phát sinh so với hợp đồng NCC/NTP/ĐTC
                                //oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                                //oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //oSheet.Cells[current_rownum, 1] = "B";
                                //oSheet.Cells[current_rownum, 2] = "KHỐI LƯỢNG PHÁT SINH SO VỚI HỢP ĐỒNG NCC/NTP/ĐTC";
                                //Section_RowNum.Add(current_rownum);
                                //current_rownum++;

                                //subprojectkey = "";
                                //Group_No = 0;
                                //Detail_No = 1;
                                //Group_No_RowNum.Clear();

                                //foreach (DataRow rB in B.Select("U_GoiThauKey=" + r["U_GoiThauKey"]))
                                //{
                                //    if (subprojectkey != rB["U_SubProjectKey"].ToString())
                                //    {
                                //        //Print group name
                                //        Group_No++;
                                //        Detail_No = 1;
                                //        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(198, 224, 180);
                                //        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //        oSheet.Cells[current_rownum, 1] = Group_No.ToString();
                                //        oSheet.Cells[current_rownum, 2] = rB["U_SubProjectName"];
                                //        subprojectkey = rB["U_SubProjectKey"].ToString();
                                //        //Sum thanh tien - Sum Gia tri thuc hien
                                //        int Detail_Row_Count = (int)B.Compute("Count(U_SubProjectKey)", string.Format("U_GoiThauKey={0} and U_SubProjectKey={1}", r["U_GoiThauKey"], rB["U_SubProjectKey"]));
                                //        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = "=SUM(G" + (current_rownum + 1).ToString() + ":G" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 9]).Formula = "=SUM(I" + (current_rownum + 1).ToString() + ":I" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //        Group_No_RowNum.Add(current_rownum);
                                //        current_rownum++;
                                //    }
                                //    //Print Details
                                //    oSheet.Cells[current_rownum, 1] = Group_No + "." + Detail_No;
                                //    oSheet.Cells[current_rownum, 2] = rB["U_OpenIssueRemark"];
                                //    oSheet.Cells[current_rownum, 4] = rB["U_UoM"];
                                //    oSheet.Cells[current_rownum, 5] = rB["U_Quantity"];
                                //    oSheet.Cells[current_rownum, 6] = rB["U_UPrice"];
                                //    oSheet.Cells[current_rownum, 7] = rB["U_Sum"];
                                //    oSheet.Cells[current_rownum, 8] = rB["U_CompleteRate"];
                                //    oSheet.Cells[current_rownum, 9] = rB["U_CompleteAmount"];
                                //    current_rownum++;
                                //    Detail_No++;
                                //}

                                //if (Group_No_RowNum.Count > 0)
                                //{
                                //    string cell_sum_tt = "";
                                //    string cell_sum_gtth = "";
                                //    int temp = 1;
                                //    foreach (int t in Group_No_RowNum)
                                //    {
                                //        if (temp < Group_No_RowNum.Count)
                                //        {
                                //            cell_sum_tt += "G" + t + ",";
                                //            cell_sum_gtth += "I" + t + ",";
                                //            temp++;
                                //        }

                                //        else
                                //        {
                                //            cell_sum_tt += "G" + t;
                                //            cell_sum_gtth += "I" + t;
                                //        }
                                //    }
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 7]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                //}

                                ////C - KHẤU TRỪ VẬT TƯ - MÁY MÓC
                                //oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                                //oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //oSheet.Cells[current_rownum, 1] = "C";
                                //oSheet.Cells[current_rownum, 2] = "KHẤU TRỪ VẬT TƯ - MÁY MÓC";
                                //Section_RowNum.Add(current_rownum);
                                //current_rownum++;

                                //subprojectkey = "";
                                //Group_No = 0;
                                //Detail_No = 1;
                                //Group_No_RowNum.Clear();

                                //foreach (DataRow rC in C.Select("U_GoiThauKey=" + r["U_GoiThauKey"]))
                                //{
                                //    //if (subprojectkey != rC["U_SubProjectKey"].ToString())
                                //    //{
                                //    //    //Print group name
                                //    //    Group_No++;
                                //    //    Detail_No = 1;
                                //    //    oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(198, 224, 180);
                                //    //    oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //    //    oSheet.Cells[current_rownum, 1] = Group_No.ToString();
                                //    //    oSheet.Cells[current_rownum, 2] = rC["U_SubProjectName"];
                                //    //    subprojectkey = rC["U_SubProjectKey"].ToString();
                                //    //    //Sum thanh tien - Sum Gia tri thuc hien
                                //    //    int Detail_Row_Count = (int)C.Compute("Count(U_SubProjectKey)", string.Format("U_GoiThauKey={0} and U_SubProjectKey={1}", r["U_GoiThauKey"], rC["U_SubProjectKey"]));
                                //    //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = "=SUM(G" + (current_rownum + 1).ToString() + ":G" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //    //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 9]).Formula = "=SUM(I" + (current_rownum + 1).ToString() + ":I" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //    //    current_rownum++;
                                //    //}
                                //    //Print Details
                                //    oSheet.Cells[current_rownum, 1] = Detail_No;
                                //    oSheet.Cells[current_rownum, 2] = rC["U_DetailsName"];
                                //    oSheet.Cells[current_rownum, 4] = rC["U_UoM"];
                                //    oSheet.Cells[current_rownum, 5] = rC["U_Quantity"];
                                //    oSheet.Cells[current_rownum, 6] = rC["U_UPrice"];
                                //    oSheet.Cells[current_rownum, 7] = rC["U_Sum"];
                                //    oSheet.Cells[current_rownum, 8] = rC["U_CompleteRate"];
                                //    oSheet.Cells[current_rownum, 9] = rC["U_CompleteAmount"];
                                //    Group_No_RowNum.Add(current_rownum);
                                //    current_rownum++;
                                //    Detail_No++;
                                //}

                                //if (Group_No_RowNum.Count > 0)
                                //{
                                //    string cell_sum_tt = "";
                                //    string cell_sum_gtth = "";
                                //    int temp = 1;
                                //    foreach (int t in Group_No_RowNum)
                                //    {
                                //        if (temp < Group_No_RowNum.Count)
                                //        {
                                //            cell_sum_tt += "G" + t + ",";
                                //            cell_sum_gtth += "I" + t + ",";
                                //            temp++;
                                //        }

                                //        else
                                //        {
                                //            cell_sum_tt += "G" + t;
                                //            cell_sum_gtth += "I" + t;
                                //        }
                                //    }
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 7]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                //}

                                ////D - KHẤU TRỪ BẢO HỘ LAO ĐỘNG
                                //oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                                //oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //oSheet.Cells[current_rownum, 1] = "D";
                                //oSheet.Cells[current_rownum, 2] = "KHẤU TRỪ BẢO HỘ LAO ĐỘNG";
                                //Section_RowNum.Add(current_rownum);
                                //current_rownum++;

                                //subprojectkey = "";
                                //Group_No = 0;
                                //Detail_No = 1;
                                //Group_No_RowNum.Clear();

                                //foreach (DataRow rD in D.Select("U_GoiThauKey=" + r["U_GoiThauKey"]))
                                //{
                                //    //if (subprojectkey != rC["U_SubProjectKey"].ToString())
                                //    //{
                                //    //    //Print group name
                                //    //    Group_No++;
                                //    //    Detail_No = 1;
                                //    //    oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(198, 224, 180);
                                //    //    oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //    //    oSheet.Cells[current_rownum, 1] = Group_No.ToString();
                                //    //    oSheet.Cells[current_rownum, 2] = rC["U_SubProjectName"];
                                //    //    subprojectkey = rC["U_SubProjectKey"].ToString();
                                //    //    //Sum thanh tien - Sum Gia tri thuc hien
                                //    //    int Detail_Row_Count = (int)C.Compute("Count(U_SubProjectKey)", string.Format("U_GoiThauKey={0} and U_SubProjectKey={1}", r["U_GoiThauKey"], rC["U_SubProjectKey"]));
                                //    //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = "=SUM(G" + (current_rownum + 1).ToString() + ":G" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //    //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 9]).Formula = "=SUM(I" + (current_rownum + 1).ToString() + ":I" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //    //    current_rownum++;
                                //    //}
                                //    //Print Details
                                //    oSheet.Cells[current_rownum, 1] = Detail_No;
                                //    oSheet.Cells[current_rownum, 2] = rD["U_DetailsName"];
                                //    oSheet.Cells[current_rownum, 4] = rD["U_UoM"];
                                //    oSheet.Cells[current_rownum, 5] = rD["U_Quantity"];
                                //    oSheet.Cells[current_rownum, 6] = rD["U_UPrice"];
                                //    oSheet.Cells[current_rownum, 7] = rD["U_Sum"];
                                //    oSheet.Cells[current_rownum, 8] = rD["U_CompleteRate"];
                                //    oSheet.Cells[current_rownum, 9] = rD["U_CompleteAmount"];
                                //    Group_No_RowNum.Add(current_rownum);
                                //    current_rownum++;
                                //    Detail_No++;
                                //}

                                //if (Group_No_RowNum.Count > 0)
                                //{
                                //    string cell_sum_tt = "";
                                //    string cell_sum_gtth = "";
                                //    int temp = 1;
                                //    foreach (int t in Group_No_RowNum)
                                //    {
                                //        if (temp < Group_No_RowNum.Count)
                                //        {
                                //            cell_sum_tt += "G" + t + ",";
                                //            cell_sum_gtth += "I" + t + ",";
                                //            temp++;
                                //        }

                                //        else
                                //        {
                                //            cell_sum_tt += "G" + t;
                                //            cell_sum_gtth += "I" + t;
                                //        }
                                //    }
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 7]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                //}

                                ////E - HỖ TRỢ THI CÔNG THEO QUY CHẾ
                                //oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                                //oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //oSheet.Cells[current_rownum, 1] = "E";
                                //oSheet.Cells[current_rownum, 2] = "HỖ TRỢ THI CÔNG THEO QUY CHẾ";
                                //Section_RowNum.Add(current_rownum);
                                //current_rownum++;

                                //subprojectkey = "";
                                //Group_No = 0;
                                //Detail_No = 1;
                                //Group_No_RowNum.Clear();

                                //foreach (DataRow rE in E.Select("U_GoiThauKey=" + r["U_GoiThauKey"]))
                                //{
                                //    //if (subprojectkey != rC["U_SubProjectKey"].ToString())
                                //    //{
                                //    //    //Print group name
                                //    //    Group_No++;
                                //    //    Detail_No = 1;
                                //    //    oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(198, 224, 180);
                                //    //    oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //    //    oSheet.Cells[current_rownum, 1] = Group_No.ToString();
                                //    //    oSheet.Cells[current_rownum, 2] = rC["U_SubProjectName"];
                                //    //    subprojectkey = rC["U_SubProjectKey"].ToString();
                                //    //    //Sum thanh tien - Sum Gia tri thuc hien
                                //    //    int Detail_Row_Count = (int)C.Compute("Count(U_SubProjectKey)", string.Format("U_GoiThauKey={0} and U_SubProjectKey={1}", r["U_GoiThauKey"], rC["U_SubProjectKey"]));
                                //    //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = "=SUM(G" + (current_rownum + 1).ToString() + ":G" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //    //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 9]).Formula = "=SUM(I" + (current_rownum + 1).ToString() + ":I" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //    //    current_rownum++;
                                //    //}
                                //    //Print Details
                                //    oSheet.Cells[current_rownum, 1] = Detail_No;
                                //    oSheet.Cells[current_rownum, 2] = rE["U_OpenIssueRemark"];
                                //    oSheet.Cells[current_rownum, 4] = rE["U_UoM"];
                                //    oSheet.Cells[current_rownum, 5] = rE["U_Quantity"];
                                //    oSheet.Cells[current_rownum, 6] = rE["U_UPrice"];
                                //    oSheet.Cells[current_rownum, 7] = rE["U_Sum"];
                                //    oSheet.Cells[current_rownum, 8] = rE["U_CompleteRate"];
                                //    oSheet.Cells[current_rownum, 9] = rE["U_CompleteAmount"];
                                //    Group_No_RowNum.Add(current_rownum);
                                //    current_rownum++;
                                //    Detail_No++;
                                //}

                                //if (Group_No_RowNum.Count > 0)
                                //{
                                //    string cell_sum_tt = "";
                                //    string cell_sum_gtth = "";
                                //    int temp = 1;
                                //    foreach (int t in Group_No_RowNum)
                                //    {
                                //        if (temp < Group_No_RowNum.Count)
                                //        {
                                //            cell_sum_tt += "G" + t + ",";
                                //            cell_sum_gtth += "I" + t + ",";
                                //            temp++;
                                //        }

                                //        else
                                //        {
                                //            cell_sum_tt += "G" + t;
                                //            cell_sum_gtth += "I" + t;
                                //        }
                                //    }
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 7]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                //}

                                ////F - HỖ TRỢ NGOÀI QUY CHẾ TRONG QUÁ TRÌNH THI CÔNG
                                //oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                                //oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //oSheet.Cells[current_rownum, 1] = "F";
                                //oSheet.Cells[current_rownum, 2] = "HỖ TRỢ NGOÀI QUY CHẾ TRONG QUÁ TRÌNH THI CÔNG";
                                //Section_RowNum.Add(current_rownum);
                                //current_rownum++;

                                //subprojectkey = "";
                                //Group_No = 0;
                                //Detail_No = 1;
                                //Group_No_RowNum.Clear();

                                //foreach (DataRow rF in F.Select("U_GoiThauKey=" + r["U_GoiThauKey"]))
                                //{
                                //    //if (subprojectkey != rC["U_SubProjectKey"].ToString())
                                //    //{
                                //    //    //Print group name
                                //    //    Group_No++;
                                //    //    Detail_No = 1;
                                //    //    oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(198, 224, 180);
                                //    //    oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //    //    oSheet.Cells[current_rownum, 1] = Group_No.ToString();
                                //    //    oSheet.Cells[current_rownum, 2] = rC["U_SubProjectName"];
                                //    //    subprojectkey = rC["U_SubProjectKey"].ToString();
                                //    //    //Sum thanh tien - Sum Gia tri thuc hien
                                //    //    int Detail_Row_Count = (int)C.Compute("Count(U_SubProjectKey)", string.Format("U_GoiThauKey={0} and U_SubProjectKey={1}", r["U_GoiThauKey"], rC["U_SubProjectKey"]));
                                //    //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = "=SUM(G" + (current_rownum + 1).ToString() + ":G" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //    //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 9]).Formula = "=SUM(I" + (current_rownum + 1).ToString() + ":I" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //    //    current_rownum++;
                                //    //}
                                //    //Print Details
                                //    oSheet.Cells[current_rownum, 1] = Detail_No;
                                //    oSheet.Cells[current_rownum, 2] = rF["U_OpenIssueRemark"];
                                //    oSheet.Cells[current_rownum, 4] = rF["U_UoM"];
                                //    oSheet.Cells[current_rownum, 5] = rF["U_Quantity"];
                                //    oSheet.Cells[current_rownum, 6] = rF["U_UPrice"];
                                //    oSheet.Cells[current_rownum, 7] = rF["U_Sum"];
                                //    oSheet.Cells[current_rownum, 8] = rF["U_CompleteRate"];
                                //    oSheet.Cells[current_rownum, 9] = rF["U_CompleteAmount"];
                                //    Group_No_RowNum.Add(current_rownum);
                                //    current_rownum++;
                                //    Detail_No++;
                                //}

                                //if (Group_No_RowNum.Count > 0)
                                //{
                                //    string cell_sum_tt = "";
                                //    string cell_sum_gtth = "";
                                //    int temp = 1;
                                //    foreach (int t in Group_No_RowNum)
                                //    {
                                //        if (temp < Group_No_RowNum.Count)
                                //        {
                                //            cell_sum_tt += "G" + t + ",";
                                //            cell_sum_gtth += "I" + t + ",";
                                //            temp++;
                                //        }

                                //        else
                                //        {
                                //            cell_sum_tt += "G" + t;
                                //            cell_sum_gtth += "I" + t;
                                //        }
                                //    }
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 7]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                //}

                                ////G - THƯỞNG, PHẠT THI CÔNG
                                //oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                                //oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //oSheet.Cells[current_rownum, 1] = "G";
                                //oSheet.Cells[current_rownum, 2] = "THƯỞNG, PHẠT THI CÔNG";
                                //Section_RowNum.Add(current_rownum);
                                //current_rownum++;

                                //subprojectkey = "";
                                //Group_No = 0;
                                //Detail_No = 1;
                                //Group_No_RowNum.Clear();

                                //foreach (DataRow rG in G.Select("U_GoiThauKey=" + r["U_GoiThauKey"]))
                                //{
                                //    //if (subprojectkey != rC["U_SubProjectKey"].ToString())
                                //    //{
                                //    //    //Print group name
                                //    //    Group_No++;
                                //    //    Detail_No = 1;
                                //    //    oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(198, 224, 180);
                                //    //    oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                //    //    oSheet.Cells[current_rownum, 1] = Group_No.ToString();
                                //    //    oSheet.Cells[current_rownum, 2] = rC["U_SubProjectName"];
                                //    //    subprojectkey = rC["U_SubProjectKey"].ToString();
                                //    //    //Sum thanh tien - Sum Gia tri thuc hien
                                //    //    int Detail_Row_Count = (int)C.Compute("Count(U_SubProjectKey)", string.Format("U_GoiThauKey={0} and U_SubProjectKey={1}", r["U_GoiThauKey"], rC["U_SubProjectKey"]));
                                //    //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = "=SUM(G" + (current_rownum + 1).ToString() + ":G" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //    //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 9]).Formula = "=SUM(I" + (current_rownum + 1).ToString() + ":I" + (current_rownum + Detail_Row_Count).ToString() + ")";
                                //    //    current_rownum++;
                                //    //}
                                //    //Print Details
                                //    oSheet.Cells[current_rownum, 1] = Detail_No;
                                //    oSheet.Cells[current_rownum, 2] = rG["U_OpenIssueRemark"];
                                //    oSheet.Cells[current_rownum, 4] = rG["U_UoM"];
                                //    oSheet.Cells[current_rownum, 5] = rG["U_Quantity"];
                                //    oSheet.Cells[current_rownum, 6] = rG["U_UPrice"];
                                //    oSheet.Cells[current_rownum, 7] = rG["U_Sum"];
                                //    oSheet.Cells[current_rownum, 8] = rG["U_CompleteRate"];
                                //    oSheet.Cells[current_rownum, 9] = rG["U_CompleteAmount"];
                                //    Group_No_RowNum.Add(current_rownum);
                                //    current_rownum++;
                                //    Detail_No++;
                                //}

                                //if (Group_No_RowNum.Count > 0)
                                //{
                                //    string cell_sum_tt = "";
                                //    string cell_sum_gtth = "";
                                //    int temp = 1;
                                //    foreach (int t in Group_No_RowNum)
                                //    {
                                //        if (temp < Group_No_RowNum.Count)
                                //        {
                                //            cell_sum_tt += "G" + t + ",";
                                //            cell_sum_gtth += "I" + t + ",";
                                //            temp++;
                                //        }

                                //        else
                                //        {
                                //            cell_sum_tt += "G" + t;
                                //            cell_sum_gtth += "I" + t;
                                //        }
                                //    }
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 7]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                //}

                                ////Sum THI CONG HOAN THIEN
                                //if (Section_RowNum.Count > 0)
                                //{
                                //    string cell_sum_tt = "";
                                //    string cell_sum_gtth = "";
                                //    int temp = 1;
                                //    foreach (int t in Section_RowNum)
                                //    {
                                //        if (temp < Section_RowNum.Count)
                                //        {
                                //            cell_sum_tt += "G" + t + ",";
                                //            cell_sum_gtth += "I" + t + ",";
                                //            temp++;
                                //        }
                                //        else
                                //        {
                                //            cell_sum_tt += "G" + t;
                                //            cell_sum_gtth += "I" + t;
                                //        }
                                //    }
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[0] - 1, 7]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[0] - 1, 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                //}
                                //Goithau_RowNum = current_rownum;
                                #endregion

                            }
                        #endregion

                            //Total A
                            if (Group_No_RowNum.Count > 0)
                            {
                                //string cell_sum_tt = "";
                                //string cell_sum_gtth = "";
                                //string cell_sum_gtth2 = "";
                                //int temp = 1;
                                //foreach (int t in Group_No_RowNum)
                                //{
                                //    if (temp < Group_No_RowNum.Count)
                                //    {
                                //        cell_sum_tt += "E" + t + ",";
                                //        cell_sum_gtth += "G" + t + ",";
                                //        cell_sum_gtth2 += "I" + t + ",";
                                //        temp++;
                                //    }

                                //    else
                                //    {
                                //        cell_sum_tt += "E" + t;
                                //        cell_sum_gtth += "G" + t;
                                //        cell_sum_gtth2 += "I" + t;
                                //    }
                                //}
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_No_RowNum[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_No_RowNum[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Group_No_RowNum[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 8]).Formula = string.Format(@"=I{0}/G{1}", Section_RowNum[(Section_RowNum.Count - 1)], Section_RowNum[(Section_RowNum.Count - 1)]);
                            }
                        }
                        #endregion

                        //B - Khối lượng phát sinh so với hợp đồng NCC/NTP/ĐTC
                        #region B
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Cells[current_rownum, 1] = "B";
                        oSheet.Cells[current_rownum, 2] = "KHỐI LƯỢNG PHÁT SINH SO VỚI HỢP ĐỒNG NCC/NTP/ĐTC";
                        Section_RowNum.Add(current_rownum);
                        current_rownum++;

                        tb_goithau = B.AsDataView().ToTable(true, new string[] { "U_Sub1", "U_Sub1Name" });
                        Group_No_RowNum.Clear();
                        Group_No_RowNum2.Clear();
                        Group_No_RowNum3.Clear();
                        Group_No_RowNum4.Clear();
                        Group_No_RowNum5.Clear();
                        #region LV1
                        if (tb_goithau.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_goithau.Rows[0]["U_Sub1"].ToString()))
                        {
                            foreach (DataRow r in tb_goithau.Rows)
                            {
                                Group_No_RowNum2.Clear();
                                Group_No_RowNum3.Clear();
                                Group_No_RowNum4.Clear();
                                Group_No_RowNum5.Clear();
                                //Goi Thau
                                List<int> Gr_element = new List<int>();
                                oSheet.Cells[current_rownum, 1] = Group_No_RowNum.Count + 1;
                                oSheet.Cells[current_rownum, 2] = r["U_Sub1Name"].ToString();
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 8]).Formula = string.Format(@"=I{0}/G{1}", current_rownum, current_rownum);
                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                Group_No_RowNum.Add(current_rownum);
                                current_rownum++;

                                //Gang chung tu o Sub1 - Goi thau
                                foreach (DataRow t in B.Select("U_Sub1 ='" + r["U_Sub1"] + "' and ISNULL(U_Sub2,'') =''"))
                                {
                                    oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                    oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                    oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                    oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                    oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                    oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                    oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                    oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                    Gr_element.Add(current_rownum);
                                    current_rownum++;
                                }

                                //Group level 2
                                #region LV2
                                DataTable tb_lv2 = Get_List_Sub_Level(int.Parse(docnum), int.Parse(r["U_Sub1"].ToString()), 2, "B");
                                if (tb_lv2.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_lv2.Rows[0]["U_Sub2"].ToString()))
                                {
                                    foreach (DataRow r_lv2 in tb_lv2.Rows)
                                    {
                                        List<int> Gr2_element = new List<int>();
                                        if (!String.IsNullOrEmpty(r_lv2["U_Sub2"].ToString()))
                                        {
                                            oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(255, 217, 102);
                                            oSheet.Cells[current_rownum, 1] = string.Format("{0}.{1}", Group_No_RowNum.Count, Group_No_RowNum2.Count + 1);// (Group_No_RowNum.Count + 1);
                                            oSheet.Cells[current_rownum, 2] = r_lv2["U_Sub2Name"].ToString();
                                            Group_No_RowNum2.Add(current_rownum);
                                            Gr_element.Add(current_rownum);
                                            current_rownum++;

                                            foreach (DataRow t in B.Select("U_Sub2 ='" + r_lv2["U_Sub2"] + "' and ISNULL(U_Sub3,'') =''"))
                                            {
                                                oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                                oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                                oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                                oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                                oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                                oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                                oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                                oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                                Gr2_element.Add(current_rownum);
                                                current_rownum++;
                                            }

                                            //Group level 3
                                            #region LV3
                                            DataTable tb_lv3 = Get_List_Sub_Level(int.Parse(docnum), int.Parse(r_lv2["U_Sub2"].ToString()), 3, "B");
                                            if (tb_lv3.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_lv3.Rows[0]["U_Sub3"].ToString()))
                                            {
                                                foreach (DataRow r_lv3 in tb_lv3.Rows)
                                                {
                                                    List<int> Gr3_element = new List<int>();
                                                    if (!String.IsNullOrEmpty(r_lv3["U_Sub3"].ToString()))
                                                    {
                                                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(255, 242, 204);
                                                        oSheet.Cells[current_rownum, 2] = r_lv3["U_Sub3Name"].ToString();
                                                        oSheet.Cells[current_rownum, 1] = string.Format("{0}.{1}.{2}", Group_No_RowNum.Count, Group_No_RowNum2.Count, Group_No_RowNum3.Count + 1);
                                                        Gr2_element.Add(current_rownum);
                                                        Group_No_RowNum3.Add(current_rownum);
                                                        current_rownum++;

                                                        foreach (DataRow t in B.Select("U_Sub3 ='" + r_lv3["U_Sub3"] + "' and ISNULL(U_Sub4,'') =''"))
                                                        {
                                                            oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                                            oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                                            oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                                            oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                                            oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                                            oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                                            oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                                            oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                                            Gr3_element.Add(current_rownum);
                                                            current_rownum++;
                                                        }

                                                        //Group level 4
                                                        #region LV4
                                                        DataTable tb_lv4 = Get_List_Sub_Level(int.Parse(docnum), int.Parse(r_lv3["U_Sub3"].ToString()), 4, "B");
                                                        if (tb_lv4.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_lv4.Rows[0]["U_Sub4"].ToString()))
                                                        {

                                                            foreach (DataRow r_lv4 in tb_lv4.Rows)
                                                            {
                                                                List<int> Gr4_element = new List<int>();
                                                                if (!String.IsNullOrEmpty(r_lv4["U_Sub4"].ToString()))
                                                                {
                                                                    oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);
                                                                    oSheet.Cells[current_rownum, 2] = r_lv4["U_Sub4Name"].ToString();
                                                                    Group_No_RowNum4.Add(current_rownum);
                                                                    Gr3_element.Add(current_rownum);
                                                                    current_rownum++;

                                                                    foreach (DataRow t in B.Select("U_Sub4 ='" + r_lv4["U_Sub4"] + "' and ISNULL(U_Sub5,'') =''"))
                                                                    {
                                                                        oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                                                        oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                                                        oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                                                        oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                                                        oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                                                        oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                                                        oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                                                        oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                                                        Gr4_element.Add(current_rownum);
                                                                        current_rownum++;

                                                                    }
                                                                    //Group level 5
                                                                    #region LV5
                                                                    DataTable tb_lv5 = Get_List_Sub_Level(int.Parse(docnum), int.Parse(r_lv4["U_Sub4"].ToString()), 5, "B");
                                                                    if (tb_lv5.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_lv5.Rows[0]["U_Sub5"].ToString()))
                                                                    {
                                                                        foreach (DataRow r_lv5 in tb_lv5.Rows)
                                                                        {
                                                                            if (!String.IsNullOrEmpty(r_lv5["U_Sub5"].ToString()))
                                                                            {
                                                                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(214, 220, 218);
                                                                                oSheet.Cells[current_rownum, 2] = r_lv5["U_Sub5Name"].ToString();
                                                                                Gr4_element.Add(current_rownum);
                                                                                Group_No_RowNum5.Add(current_rownum);
                                                                                current_rownum++;

                                                                                //Detail
                                                                                foreach (DataRow t in B.Select("U_Sub5 ='" + r_lv5["U_Sub5"] + "'"))
                                                                                {
                                                                                    oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                                                                    oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                                                                    oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                                                                    oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                                                                    oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                                                                    oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                                                                    oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                                                                    oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                                                                    current_rownum++;
                                                                                }
                                                                            }
                                                                        }
                                                                        //Total Level 5
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1, current_rownum - 1); // "=SUM(E" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":E" + (current_rownum - 1).ToString() + ")";
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1, current_rownum - 1); //"=SUM(G" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":G" + (current_rownum - 1).ToString() + ")";
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1, current_rownum - 1); // "=SUM(I" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":I" + (current_rownum - 1).ToString() + ")";
                                                                    }

                                                                    #endregion

                                                                    //Total Level 4
                                                                    if (Gr4_element.Count > 0)
                                                                    {
                                                                        //string cell_sum_tt = "";
                                                                        //string cell_sum_gtth = "";
                                                                        //string cell_sum_gtth2 = "";
                                                                        //int temp = 1;
                                                                        //foreach (int t in Gr4_element)
                                                                        //{
                                                                        //    if (temp < Gr4_element.Count)
                                                                        //    {
                                                                        //        cell_sum_tt += "E" + t + ",";
                                                                        //        cell_sum_gtth += "G" + t + ",";
                                                                        //        cell_sum_gtth2 += "I" + t + ",";
                                                                        //        temp++;
                                                                        //    }

                                                                        //    else
                                                                        //    {
                                                                        //        cell_sum_tt += "E" + t;
                                                                        //        cell_sum_gtth += "G" + t;
                                                                        //        cell_sum_gtth2 += "I" + t;
                                                                        //    }
                                                                        //}
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr4_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Gr4_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Gr4_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                                                    }
                                                                }

                                                            }
                                                        }
                                                        #endregion

                                                        //Total level 3
                                                        if (Gr3_element.Count > 0)
                                                        {
                                                            //string cell_sum_tt = "";
                                                            //string cell_sum_gtth = "";
                                                            //string cell_sum_gtth2 = "";
                                                            //int temp = 1;
                                                            //foreach (int t in Gr3_element)
                                                            //{
                                                            //    if (temp < Gr3_element.Count)
                                                            //    {
                                                            //        cell_sum_tt += "E" + t + ",";
                                                            //        cell_sum_gtth += "G" + t + ",";
                                                            //        cell_sum_gtth2 += "I" + t + ",";
                                                            //        temp++;
                                                            //    }

                                                            //    else
                                                            //    {
                                                            //        cell_sum_tt += "E" + t;
                                                            //        cell_sum_gtth += "G" + t;
                                                            //        cell_sum_gtth2 += "I" + t;
                                                            //    }
                                                            //}
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr3_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Gr3_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Gr3_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion

                                            //Total level 2
                                            if (Gr2_element.Count > 0)
                                            {
                                                //string cell_sum_tt = "";
                                                //string cell_sum_gtth = "";
                                                //string cell_sum_gtth2 = "";
                                                //int temp = 1;
                                                //foreach (int t in Gr2_element)
                                                //{
                                                //    if (temp < Gr2_element.Count)
                                                //    {
                                                //        cell_sum_tt += "E" + t + ",";
                                                //        cell_sum_gtth += "G" + t + ",";
                                                //        cell_sum_gtth2 += "I" + t + ",";
                                                //        temp++;
                                                //    }

                                                //    else
                                                //    {
                                                //        cell_sum_tt += "E" + t;
                                                //        cell_sum_gtth += "G" + t;
                                                //        cell_sum_gtth2 += "I" + t;
                                                //    }
                                                //}
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr2_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr2_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr2_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                            }
                                            Group_No_RowNum3.Clear();
                                        }
                                    }
                                }
                                #endregion

                                //Total Goi Thau LV 1
                                if (Gr_element.Count > 0)
                                {
                                    //string cell_sum_tt = "";
                                    //string cell_sum_gtth = "";
                                    //string cell_sum_gtth2 = "";
                                    //int temp = 1;
                                    //foreach (int t in Gr_element)
                                    //{
                                    //    if (temp < Gr_element.Count)
                                    //    {
                                    //        cell_sum_tt += "E" + t + ",";
                                    //        cell_sum_gtth += "G" + t + ",";
                                    //        cell_sum_gtth2 += "I" + t + ",";
                                    //        temp++;
                                    //    }

                                    //    else
                                    //    {
                                    //        cell_sum_tt += "E" + t;
                                    //        cell_sum_gtth += "G" + t;
                                    //        cell_sum_gtth2 += "I" + t;
                                    //    }
                                    //}
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Gr_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Gr_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                }
                            }
                        #endregion

                            //Total B
                            if (Group_No_RowNum.Count > 0)
                            {
                                string cell_sum_tt = "";
                                string cell_sum_gtth = "";
                                string cell_sum_gtth2 = "";
                                int temp = 1;
                                foreach (int t in Group_No_RowNum)
                                {
                                    if (temp < Group_No_RowNum.Count)
                                    {
                                        cell_sum_tt += "E" + t + ",";
                                        cell_sum_gtth += "G" + t + ",";
                                        cell_sum_gtth2 += "I" + t + ",";
                                        temp++;
                                    }

                                    else
                                    {
                                        cell_sum_tt += "E" + t;
                                        cell_sum_gtth += "G" + t;
                                        cell_sum_gtth2 += "I" + t;
                                    }
                                }
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_No_RowNum[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_No_RowNum[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Group_No_RowNum[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 8]).Formula = string.Format(@"=I{0}/G{1}", Section_RowNum[(Section_RowNum.Count - 1)], Section_RowNum[(Section_RowNum.Count - 1)]);
                            }
                        }
                        #endregion

                        //C - KHẤU TRỪ VẬT TƯ - MÁY MÓC
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Cells[current_rownum, 1] = "C";
                        oSheet.Cells[current_rownum, 2] = "KHẤU TRỪ VẬT TƯ - MÁY MÓC";
                        Section_RowNum.Add(current_rownum);
                        current_rownum++;
                        Detail_No = 1;
                        Group_No_RowNum.Clear();
                        foreach (DataRow rC in C.Rows)
                        {
                            //Print Details
                            if (!string.IsNullOrEmpty(rC["U_DetailsName"].ToString()))
                            {
                                oSheet.Cells[current_rownum, 1] = Detail_No;
                                oSheet.Cells[current_rownum, 2] = rC["U_DetailsName"];
                                oSheet.Cells[current_rownum, 4] = rC["U_UoM"];
                                oSheet.Cells[current_rownum, 5] = rC["U_Quantity"];
                                oSheet.Cells[current_rownum, 6] = rC["U_UPrice"];
                                oSheet.Cells[current_rownum, 7] = rC["U_Sum"];
                                oSheet.Cells[current_rownum, 8] = rC["U_CompleteRate"];
                                oSheet.Cells[current_rownum, 9] = rC["U_CompleteAmount"];
                                Group_No_RowNum.Add(current_rownum);
                                current_rownum++;
                                Detail_No++;
                            }
                        }
                        //Total C
                        if (Group_No_RowNum.Count > 0)
                        {
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 8]).Formula = string.Format(@"=I{0}/G{1}", Section_RowNum[(Section_RowNum.Count - 1)], Section_RowNum[(Section_RowNum.Count - 1)]);
                        }
                        //D - KHẤU TRỪ BẢO HỘ LAO ĐỘNG
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Cells[current_rownum, 1] = "D";
                        oSheet.Cells[current_rownum, 2] = "KHẤU TRỪ BẢO HỘ LAO ĐỘNG";
                        Section_RowNum.Add(current_rownum);
                        Group_No_RowNum.Clear();
                        current_rownum++;
                        Detail_No = 1;
                        foreach (DataRow rD in D.Rows)
                        {
                            //Print Details
                            if (!string.IsNullOrEmpty(rD["U_DetailsName"].ToString()))
                            {
                                oSheet.Cells[current_rownum, 1] = Detail_No;
                                oSheet.Cells[current_rownum, 2] = rD["U_DetailsName"];
                                oSheet.Cells[current_rownum, 4] = rD["U_UoM"];
                                oSheet.Cells[current_rownum, 5] = rD["U_Quantity"];
                                oSheet.Cells[current_rownum, 6] = rD["U_UPrice"];
                                oSheet.Cells[current_rownum, 7] = rD["U_Sum"];
                                oSheet.Cells[current_rownum, 8] = rD["U_CompleteRate"];
                                oSheet.Cells[current_rownum, 9] = rD["U_CompleteAmount"];
                                Group_No_RowNum.Add(current_rownum);
                                current_rownum++;
                                Detail_No++;
                            }
                        }
                        //Total D
                        if (Group_No_RowNum.Count > 0)
                        {
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 8]).Formula = string.Format(@"=I{0}/G{1}", Section_RowNum[(Section_RowNum.Count - 1)], Section_RowNum[(Section_RowNum.Count - 1)]);
                        }
                        //E - HỖ TRỢ THI CÔNG THEO QUY CHẾ
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Cells[current_rownum, 1] = "E";
                        oSheet.Cells[current_rownum, 2] = "HỖ TRỢ THI CÔNG THEO QUY CHẾ";
                        Section_RowNum.Add(current_rownum);
                        current_rownum++;
                        Group_No_RowNum.Clear();
                        Detail_No = 1;
                        foreach (DataRow rE in E.Rows)
                        {
                            //Print Details
                            if (!string.IsNullOrEmpty(rE["U_OpenIssueKey"].ToString()))
                            {
                                oSheet.Cells[current_rownum, 1] = Detail_No;
                                oSheet.Cells[current_rownum, 2] = rE["U_OpenIssueRemark"];
                                oSheet.Cells[current_rownum, 4] = rE["U_UoM"];
                                oSheet.Cells[current_rownum, 5] = rE["U_Quantity"];
                                oSheet.Cells[current_rownum, 6] = rE["U_UPrice"];
                                oSheet.Cells[current_rownum, 7] = rE["U_Sum"];
                                oSheet.Cells[current_rownum, 8] = rE["U_CompleteRate"];
                                oSheet.Cells[current_rownum, 9] = rE["U_CompleteAmount"];
                                Group_No_RowNum.Add(current_rownum);
                                current_rownum++;
                                Detail_No++;
                            }
                        }
                        //Total E
                        if (Group_No_RowNum.Count > 0)
                        {
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 8]).Formula = string.Format(@"=I{0}/G{1}", Section_RowNum[(Section_RowNum.Count - 1)], Section_RowNum[(Section_RowNum.Count - 1)]);
                        }
                        //F - HỖ TRỢ NGOÀI QUY CHẾ TRONG QUÁ TRÌNH THI CÔNG
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Cells[current_rownum, 1] = "F";
                        oSheet.Cells[current_rownum, 2] = "HỖ TRỢ NGOÀI QUY CHẾ TRONG QUÁ TRÌNH THI CÔNG";
                        Section_RowNum.Add(current_rownum);
                        current_rownum++;
                        Detail_No = 1;
                        Group_No_RowNum.Clear();

                        foreach (DataRow rF in F.Rows)
                        {
                            //Print Details
                            if (!string.IsNullOrEmpty(rF["U_OpenIssueKey"].ToString()))
                            {
                                oSheet.Cells[current_rownum, 1] = Detail_No;
                                oSheet.Cells[current_rownum, 2] = rF["U_OpenIssueRemark"];
                                oSheet.Cells[current_rownum, 4] = rF["U_UoM"];
                                oSheet.Cells[current_rownum, 5] = rF["U_Quantity"];
                                oSheet.Cells[current_rownum, 6] = rF["U_UPrice"];
                                oSheet.Cells[current_rownum, 7] = rF["U_Sum"];
                                oSheet.Cells[current_rownum, 8] = rF["U_CompleteRate"];
                                oSheet.Cells[current_rownum, 9] = rF["U_CompleteAmount"];
                                Group_No_RowNum.Add(current_rownum);
                                current_rownum++;
                                Detail_No++;
                            }
                        }
                        //Total F
                        if (Group_No_RowNum.Count > 0)
                        {
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 8]).Formula = string.Format(@"=I{0}/G{1}", Section_RowNum[(Section_RowNum.Count - 1)], Section_RowNum[(Section_RowNum.Count - 1)]);
                        }
                        //G - THƯỞNG, PHẠT THI CÔNG
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Cells[current_rownum, 1] = "G";
                        oSheet.Cells[current_rownum, 2] = "THƯỞNG, PHẠT THI CÔNG";
                        Section_RowNum.Add(current_rownum);
                        Group_No_RowNum.Clear();
                        Detail_No = 1;
                        current_rownum++;

                        foreach (DataRow rG in G.Rows)
                        {
                            //Print Details
                            if (!string.IsNullOrEmpty(rG["U_OpenIssueKey"].ToString()))
                            {
                                oSheet.Cells[current_rownum, 1] = Detail_No;
                                oSheet.Cells[current_rownum, 2] = rG["U_OpenIssueRemark"];
                                oSheet.Cells[current_rownum, 4] = rG["U_UoM"];
                                oSheet.Cells[current_rownum, 5] = rG["U_Quantity"];
                                oSheet.Cells[current_rownum, 6] = rG["U_UPrice"];
                                oSheet.Cells[current_rownum, 7] = rG["U_Sum"];
                                oSheet.Cells[current_rownum, 8] = rG["U_CompleteRate"];
                                oSheet.Cells[current_rownum, 9] = rG["U_CompleteAmount"];
                                Group_No_RowNum.Add(current_rownum);
                                current_rownum++;
                                Detail_No++;
                            }
                        }
                        //Total G
                        if (Group_No_RowNum.Count > 0)
                        {
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 8]).Formula = string.Format(@"=I{0}/G{1}", Section_RowNum[(Section_RowNum.Count - 1)], Section_RowNum[(Section_RowNum.Count - 1)]);
                        }
                        //PHAT SINH
                        oSheet.Cells[current_rownum, 2] = "PHÁT SINH";
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 2]).Font.Bold = true;
                        Section_RowNum.Add(current_rownum);
                        current_rownum++;

                        //K - PS NEW
                        #region K
                        //oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(169, 208, 142);
                        //oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        //oSheet.Cells[current_rownum, 1] = "B";
                        //oSheet.Cells[current_rownum, 2] = "KHỐI LƯỢNG PHÁT SINH SO VỚI HỢP ĐỒNG NCC/NTP/ĐTC";
                        //current_rownum++;

                        tb_goithau = K.AsDataView().ToTable(true, new string[] { "U_Sub1", "U_Sub1Name" });
                        Group_No_RowNum.Clear();
                        Group_No_RowNum2.Clear();
                        Group_No_RowNum3.Clear();
                        Group_No_RowNum4.Clear();
                        Group_No_RowNum5.Clear();
                        #region LV1
                        if (tb_goithau.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_goithau.Rows[0]["U_Sub1"].ToString()))
                        {
                            foreach (DataRow r in tb_goithau.Rows)
                            {
                                Group_No_RowNum2.Clear();
                                Group_No_RowNum3.Clear();
                                Group_No_RowNum4.Clear();
                                Group_No_RowNum5.Clear();
                                //Goi Thau
                                List<int> Gr_element = new List<int>();
                                oSheet.Cells[current_rownum, 1] = Group_No_RowNum.Count + 1;
                                oSheet.Cells[current_rownum, 2] = r["U_Sub1Name"].ToString();
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 8]).Formula = string.Format(@"=I{0}/G{1}", current_rownum, current_rownum);
                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(244, 176, 132);
                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                                Group_No_RowNum.Add(current_rownum);
                                current_rownum++;

                                //Gang chung tu o Sub1 - Goi thau
                                foreach (DataRow t in K.Select("U_Sub1 ='" + r["U_Sub1"] + "' and ISNULL(U_Sub2,'') =''"))
                                {
                                    oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                    oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                    oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                    oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                    oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                    oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                    oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                    oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                    Gr_element.Add(current_rownum);
                                    current_rownum++;
                                }

                                //Group level 2
                                #region LV2
                                DataTable tb_lv2 = Get_List_Sub_Level(int.Parse(docnum), int.Parse(r["U_Sub1"].ToString()), 2, "K");
                                if (tb_lv2.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_lv2.Rows[0]["U_Sub2"].ToString()))
                                {
                                    foreach (DataRow r_lv2 in tb_lv2.Rows)
                                    {
                                        List<int> Gr2_element = new List<int>();
                                        if (!String.IsNullOrEmpty(r_lv2["U_Sub2"].ToString()))
                                        {
                                            oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(255, 217, 102);
                                            oSheet.Cells[current_rownum, 1] = string.Format("{0}.{1}", Group_No_RowNum.Count, Group_No_RowNum2.Count + 1);// (Group_No_RowNum.Count + 1);
                                            oSheet.Cells[current_rownum, 2] = r_lv2["U_Sub2Name"].ToString();
                                            Group_No_RowNum2.Add(current_rownum);
                                            Gr_element.Add(current_rownum);
                                            current_rownum++;

                                            foreach (DataRow t in K.Select("U_Sub2 ='" + r_lv2["U_Sub2"] + "' and ISNULL(U_Sub3,'') =''"))
                                            {
                                                oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                                oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                                oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                                oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                                oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                                oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                                oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                                oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                                Gr2_element.Add(current_rownum);
                                                current_rownum++;
                                            }

                                            //Group level 3
                                            #region LV3
                                            DataTable tb_lv3 = Get_List_Sub_Level(int.Parse(docnum), int.Parse(r_lv2["U_Sub2"].ToString()), 3, "K");
                                            if (tb_lv3.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_lv3.Rows[0]["U_Sub3"].ToString()))
                                            {
                                                foreach (DataRow r_lv3 in tb_lv3.Rows)
                                                {
                                                    List<int> Gr3_element = new List<int>();
                                                    if (!String.IsNullOrEmpty(r_lv3["U_Sub3"].ToString()))
                                                    {
                                                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(255, 242, 204);
                                                        oSheet.Cells[current_rownum, 2] = r_lv3["U_Sub3Name"].ToString();
                                                        oSheet.Cells[current_rownum, 1] = string.Format("{0}.{1}.{2}", Group_No_RowNum.Count, Group_No_RowNum2.Count, Group_No_RowNum3.Count + 1);
                                                        Gr2_element.Add(current_rownum);
                                                        Group_No_RowNum3.Add(current_rownum);
                                                        current_rownum++;

                                                        foreach (DataRow t in K.Select("U_Sub3 ='" + r_lv3["U_Sub3"] + "' and ISNULL(U_Sub4,'') =''"))
                                                        {
                                                            oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                                            oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                                            oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                                            oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                                            oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                                            oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                                            oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                                            oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                                            Gr3_element.Add(current_rownum);
                                                            current_rownum++;
                                                        }

                                                        //Group level 4
                                                        #region LV4
                                                        DataTable tb_lv4 = Get_List_Sub_Level(int.Parse(docnum), int.Parse(r_lv3["U_Sub3"].ToString()), 4, "K");
                                                        if (tb_lv4.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_lv4.Rows[0]["U_Sub4"].ToString()))
                                                        {

                                                            foreach (DataRow r_lv4 in tb_lv4.Rows)
                                                            {
                                                                List<int> Gr4_element = new List<int>();
                                                                if (!String.IsNullOrEmpty(r_lv4["U_Sub4"].ToString()))
                                                                {
                                                                    oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(155, 194, 230);
                                                                    oSheet.Cells[current_rownum, 2] = r_lv4["U_Sub4Name"].ToString();
                                                                    Group_No_RowNum4.Add(current_rownum);
                                                                    Gr3_element.Add(current_rownum);
                                                                    current_rownum++;

                                                                    foreach (DataRow t in K.Select("U_Sub4 ='" + r_lv4["U_Sub4"] + "' and ISNULL(U_Sub5,'') =''"))
                                                                    {
                                                                        oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                                                        oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                                                        oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                                                        oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                                                        oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                                                        oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                                                        oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                                                        oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                                                        Gr4_element.Add(current_rownum);
                                                                        current_rownum++;

                                                                    }
                                                                    //Group level 5
                                                                    #region LV5
                                                                    DataTable tb_lv5 = Get_List_Sub_Level(int.Parse(docnum), int.Parse(r_lv4["U_Sub4"].ToString()), 5, "K");
                                                                    if (tb_lv5.Rows.Count >= 1 && !string.IsNullOrEmpty(tb_lv5.Rows[0]["U_Sub5"].ToString()))
                                                                    {
                                                                        foreach (DataRow r_lv5 in tb_lv5.Rows)
                                                                        {
                                                                            if (!String.IsNullOrEmpty(r_lv5["U_Sub5"].ToString()))
                                                                            {
                                                                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(214, 220, 218);
                                                                                oSheet.Cells[current_rownum, 2] = r_lv5["U_Sub5Name"].ToString();
                                                                                Gr4_element.Add(current_rownum);
                                                                                Group_No_RowNum5.Add(current_rownum);
                                                                                current_rownum++;

                                                                                //Detail
                                                                                foreach (DataRow t in K.Select("U_Sub5 ='" + r_lv5["U_Sub5"] + "'"))
                                                                                {
                                                                                    oSheet.Cells[current_rownum, 2] = t["U_GPDetailsName"];
                                                                                    oSheet.Cells[current_rownum, 3] = t["U_CTCV"];
                                                                                    oSheet.Cells[current_rownum, 4] = t["U_UoM"];
                                                                                    oSheet.Cells[current_rownum, 5] = t["U_Quantity"];
                                                                                    oSheet.Cells[current_rownum, 6] = t["U_UPrice"];
                                                                                    oSheet.Cells[current_rownum, 7] = t["U_Sum"];
                                                                                    oSheet.Cells[current_rownum, 8] = t["U_CompleteRate"];
                                                                                    oSheet.Cells[current_rownum, 9] = t["U_CompleteAmount"];
                                                                                    current_rownum++;
                                                                                }
                                                                            }
                                                                        }
                                                                        //Total Level 5
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1, current_rownum - 1); //"=SUM(E" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":E" + (current_rownum - 1).ToString() + ")";
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1, current_rownum - 1); //"=SUM(G" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":G" + (current_rownum - 1).ToString() + ")";
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1, current_rownum - 1); //"=SUM(I" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":I" + (current_rownum - 1).ToString() + ")";
                                                                    }

                                                                    #endregion

                                                                    //Total Level 4
                                                                    if (Gr4_element.Count > 0)
                                                                    {
                                                                        //string cell_sum_tt = "";
                                                                        //string cell_sum_gtth = "";
                                                                        //string cell_sum_gtth2 = "";
                                                                        //int temp = 1;
                                                                        //foreach (int t in Gr4_element)
                                                                        //{
                                                                        //    if (temp < Gr4_element.Count)
                                                                        //    {
                                                                        //        cell_sum_tt += "E" + t + ",";
                                                                        //        cell_sum_gtth += "G" + t + ",";
                                                                        //        cell_sum_gtth2 += "I" + t + ",";
                                                                        //        temp++;
                                                                        //    }

                                                                        //    else
                                                                        //    {
                                                                        //        cell_sum_tt += "E" + t;
                                                                        //        cell_sum_gtth += "G" + t;
                                                                        //        cell_sum_gtth2 += "I" + t;
                                                                        //    }
                                                                        //}
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr4_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Gr4_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Gr4_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                                                    }
                                                                }

                                                            }
                                                        }
                                                        #endregion

                                                        //Total level 3
                                                        if (Gr3_element.Count > 0)
                                                        {
                                                            //string cell_sum_tt = "";
                                                            //string cell_sum_gtth = "";
                                                            //string cell_sum_gtth2 = "";
                                                            //int temp = 1;
                                                            //foreach (int t in Gr3_element)
                                                            //{
                                                            //    if (temp < Gr3_element.Count)
                                                            //    {
                                                            //        cell_sum_tt += "E" + t + ",";
                                                            //        cell_sum_gtth += "G" + t + ",";
                                                            //        cell_sum_gtth2 += "I" + t + ",";
                                                            //        temp++;
                                                            //    }

                                                            //    else
                                                            //    {
                                                            //        cell_sum_tt += "E" + t;
                                                            //        cell_sum_gtth += "G" + t;
                                                            //        cell_sum_gtth2 += "I" + t;
                                                            //    }
                                                            //}
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr3_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Gr3_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Gr3_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion

                                            //Total level 2
                                            if (Gr2_element.Count > 0)
                                            {
                                                string cell_sum_tt = "";
                                                string cell_sum_gtth = "";
                                                string cell_sum_gtth2 = "";
                                                int temp = 1;
                                                foreach (int t in Gr2_element)
                                                {
                                                    if (temp < Gr2_element.Count)
                                                    {
                                                        cell_sum_tt += "E" + t + ",";
                                                        cell_sum_gtth += "G" + t + ",";
                                                        cell_sum_gtth2 += "I" + t + ",";
                                                        temp++;
                                                    }

                                                    else
                                                    {
                                                        cell_sum_tt += "E" + t;
                                                        cell_sum_gtth += "G" + t;
                                                        cell_sum_gtth2 += "I" + t;
                                                    }
                                                }
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr2_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Gr2_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Gr2_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                            }
                                            Group_No_RowNum3.Clear();
                                        }
                                    }
                                }
                                #endregion

                                //Total Goi Thau LV 1
                                if (Gr_element.Count > 0)
                                {
                                    string cell_sum_tt = "";
                                    string cell_sum_gtth = "";
                                    string cell_sum_gtth2 = "";
                                    int temp = 1;
                                    foreach (int t in Gr_element)
                                    {
                                        if (temp < Gr_element.Count)
                                        {
                                            cell_sum_tt += "E" + t + ",";
                                            cell_sum_gtth += "G" + t + ",";
                                            cell_sum_gtth2 += "I" + t + ",";
                                            temp++;
                                        }

                                        else
                                        {
                                            cell_sum_tt += "E" + t;
                                            cell_sum_gtth += "G" + t;
                                            cell_sum_gtth2 += "I" + t;
                                        }
                                    }
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Gr_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_tt);
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Gr_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Gr_element[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                }
                            }
                        #endregion

                            //Total K
                            if (Group_No_RowNum.Count > 0)
                            {
                                string cell_sum_tt = "";
                                string cell_sum_gtth = "";
                                string cell_sum_gtth2 = "";
                                int temp = 1;
                                foreach (int t in Group_No_RowNum)
                                {
                                    if (temp < Group_No_RowNum.Count)
                                    {
                                        cell_sum_tt += "E" + t + ",";
                                        cell_sum_gtth += "G" + t + ",";
                                        cell_sum_gtth2 += "I" + t + ",";
                                        temp++;
                                    }

                                    else
                                    {
                                        cell_sum_tt += "E" + t;
                                        cell_sum_gtth += "G" + t;
                                        cell_sum_gtth2 += "I" + t;
                                    }
                                }
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUBTOTAL(9,E{0}:E{1})", Group_No_RowNum[0], current_rownum - 1); // string.Format("=SUM({0})", cell_sum_tt);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Group_No_RowNum[0], current_rownum - 1); // string.Format("=SUM({0})", cell_sum_gtth);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Group_No_RowNum[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth2);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 8]).Formula = string.Format(@"=I{0}/G{1}", Section_RowNum[(Section_RowNum.Count - 1)], Section_RowNum[(Section_RowNum.Count - 1)]);
                            }
                        }
                        #endregion

                        #region Old PS
//                        //A
//                        oSheet.Cells[current_rownum, 1] = "A";
//                        oSheet.Cells[current_rownum, 2] = "Khối lượng phát sinh đã được CĐT duyệt";
//                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
//                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
//                        //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 1]).Font.Bold = true;
//                        //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 2]).Font.Bold = true;
//                        Section_RowNum.Add(current_rownum);
//                        current_rownum++;

//                        subprojectkey = "";
//                        Group_No = 0;
//                        Detail_No = 1;
//                        Group_No_RowNum.Clear();

//                        foreach (DataRow rH in H.Select("U_Type='A'"))
//                        {
//                            if (subprojectkey != rH["U_PBAKey"].ToString())
//                            {
//                                //Print group name
//                                Group_No++;
//                                Detail_No = 1;
//                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(198, 224, 180);
//                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
//                                oSheet.Cells[current_rownum, 1] = Group_No.ToString();
//                                oSheet.Cells[current_rownum, 2] = string.Format("PLHĐ số: {0}, ngày: {1}", rH["U_PBANumber"].ToString(), DateTime.Parse(rH["U_PBADate"].ToString()).ToString("dd/MM/yyyy")); //rH["U_SubProjectName"];
//                                subprojectkey = rH["U_PBAKey"].ToString();
//                                //Sum thanh tien - Sum Gia tri thuc hien
//                                int Detail_Row_Count = (int)H.Compute("Count(U_PBAKey)", string.Format("U_PBAKey={0} and U_Type='A'", rH["U_PBAKey"]));
//                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = "=SUM(G" + (current_rownum + 1).ToString() + ":G" + (current_rownum + Detail_Row_Count).ToString() + ")";
//                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 9]).Formula = "=SUM(I" + (current_rownum + 1).ToString() + ":I" + (current_rownum + Detail_Row_Count).ToString() + ")";
//                                Group_No_RowNum.Add(current_rownum);
//                                current_rownum++;
//                            }
//                            //Print Details
//                            oSheet.Cells[current_rownum, 1] = Group_No + "." + Detail_No;
//                            oSheet.Cells[current_rownum, 2] = rH["U_ItemName"];
//                            oSheet.Cells[current_rownum, 4] = rH["U_UoM"];
//                            oSheet.Cells[current_rownum, 5] = rH["U_Quantity"];
//                            oSheet.Cells[current_rownum, 6] = rH["U_UPrice"];
//                            oSheet.Cells[current_rownum, 7] = rH["U_Sum"];
//                            oSheet.Cells[current_rownum, 9] = rH["U_Sum"];

//                            current_rownum++;
//                            Detail_No++;
//                        }

//                        if (Group_No_RowNum.Count > 0)
//                        {
//                            string cell_sum_tt = "";
//                            string cell_sum_gtth = "";
//                            int temp = 1;
//                            foreach (int t in Group_No_RowNum)
//                            {
//                                if (temp < Group_No_RowNum.Count)
//                                {
//                                    cell_sum_tt += "G" + t + ",";
//                                    cell_sum_gtth += "I" + t + ",";
//                                    temp++;
//                                }

//                                else
//                                {
//                                    cell_sum_tt += "G" + t;
//                                    cell_sum_gtth += "I" + t;
//                                }
//                            }
//                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 7]).Formula = string.Format("=SUM({0})", cell_sum_tt);
//                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
//                        }
//                        //B
//                        oSheet.Cells[current_rownum, 1] = "B";
//                        oSheet.Cells[current_rownum, 2] = "Khối lượng phát sinh chờ TV/CĐT duyệt";
//                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
//                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
//                        Section_RowNum.Add(current_rownum);
//                        current_rownum++;

//                        subprojectkey = "";
//                        Group_No = 0;
//                        Detail_No = 1;
//                        Group_No_RowNum.Clear();

//                        foreach (DataRow rH in H.Select("U_Type='D'"))
//                        {
//                            if (subprojectkey != rH["U_PBAKey"].ToString())
//                            {
//                                //Print group name
//                                Group_No++;
//                                Detail_No = 1;
//                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(198, 224, 180);
//                                oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
//                                oSheet.Cells[current_rownum, 1] = Group_No.ToString();
//                                oSheet.Cells[current_rownum, 2] = string.Format("PLHĐ số: {0}, ngày: {1}", rH["U_PBANumber"].ToString(), DateTime.Parse(rH["U_PBADate"].ToString()).ToString("dd/MM/yyyy")); //rH["U_SubProjectName"];
//                                subprojectkey = rH["U_PBAKey"].ToString();
//                                //Sum thanh tien - Sum Gia tri thuc hien
//                                int Detail_Row_Count = (int)H.Compute("Count(U_PBAKey)", string.Format("U_PBAKey={0} and U_Type='D'", rH["U_PBAKey"]));
//                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = "=SUM(G" + (current_rownum + 1).ToString() + ":G" + (current_rownum + Detail_Row_Count).ToString() + ")";
//                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 9]).Formula = "=SUM(I" + (current_rownum + 1).ToString() + ":I" + (current_rownum + Detail_Row_Count).ToString() + ")";
//                                Group_No_RowNum.Add(current_rownum);
//                                current_rownum++;
//                            }
//                            //Print Details
//                            oSheet.Cells[current_rownum, 1] = Group_No + "." + Detail_No;
//                            oSheet.Cells[current_rownum, 2] = rH["U_ItemName"];
//                            oSheet.Cells[current_rownum, 4] = rH["U_UoM"];
//                            oSheet.Cells[current_rownum, 5] = rH["U_Quantity"];
//                            oSheet.Cells[current_rownum, 6] = rH["U_UPrice"];
//                            oSheet.Cells[current_rownum, 7] = rH["U_Sum"];
//                            oSheet.Cells[current_rownum, 9] = rH["U_Sum"];
//                            current_rownum++;
//                            Detail_No++;
//                        }

//                        if (Group_No_RowNum.Count > 0)
//                        {
//                            string cell_sum_tt = "";
//                            string cell_sum_gtth = "";
//                            int temp = 1;
//                            foreach (int t in Group_No_RowNum)
//                            {
//                                if (temp < Group_No_RowNum.Count)
//                                {
//                                    cell_sum_tt += "G" + t + ",";
//                                    cell_sum_gtth += "I" + t + ",";
//                                    temp++;
//                                }

//                                else
//                                {
//                                    cell_sum_tt += "G" + t;
//                                    cell_sum_gtth += "I" + t;
//                                }
//                            }
//                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 7]).Formula = string.Format("=SUM({0})", cell_sum_tt);
//                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[0] - 1, 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
//                        }
#endregion

                        //Chi phi QL - DTC Nhan tri
                        if (BPGCode == "112")
                        {
                            oSheet.Cells[current_rownum, 2] = "CHI PHÍ QUẢN LÝ";
                            oSheet.Cells[current_rownum, 4] = "%";
                            oSheet.Cells[current_rownum, 5].Formula = string.Format("={0}/100", PhiQL);
                            oSheet.Range["E" + current_rownum].NumberFormat = "0.00%";
                            oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 2]).Font.Bold = true;
                            //Section_RowNum.Add(current_rownum);
                            current_rownum++;
                        }

                        //TOTAL
                        int Tong_GT_RowNum = current_rownum;
                        oSheet.Cells[current_rownum, 2] = "TỔNG GIÁ TRỊ (Chưa VAT)";
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(226, 239, 218);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        //string cell_sum_total = "";
                        if (Section_RowNum.Count > 0)
                        {
                            //string cell_sum_gtth = "";
                            //int temp = 1;
                            //foreach (int t in Section_RowNum)
                            //{
                            //    if (temp < Section_RowNum.Count)
                            //    {
                            //        cell_sum_gtth += "I" + t + ",";
                            //        cell_sum_total += "G" + t + ",";
                            //        temp++;
                            //    }
                            //    else
                            //    {
                            //        cell_sum_gtth += "I" + t;
                            //        cell_sum_total += "G" + t;
                            //    }
                            //}
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 9]).Formula = string.Format("=SUBTOTAL(9,I{0}:I{1})", Section_RowNum[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_gtth);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = string.Format("=SUBTOTAL(9,G{0}:G{1})", Section_RowNum[0], current_rownum - 1); //string.Format("=SUM({0})", cell_sum_total);
                        }

                        //Tinh gia tri Chi phi QL Nhan Tri Tu dong Total
                        if (BPGCode == "112")
                        {
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum - 1, 7]).Formula = string.Format("={0}*E{1}", oSheet.Cells[current_rownum, 7].Value2, current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum - 1, 9]).Formula = string.Format("={0}*E{1}", oSheet.Cells[current_rownum, 9].Value2, current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum - 1, 8]).Formula = string.Format("=I{0}/G{0}", current_rownum - 1);
                        }
                        current_rownum++;

                        int VAT_Rownum = 0;
                        oSheet.Cells[current_rownum, 2] = "% thuế VAT";
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(226, 239, 218);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        float percent = 0;
                        float.TryParse(dt_tmp.Rows[0]["U_VAT"].ToString(), out percent);
                        oSheet.Cells[current_rownum, 8] = (percent / 100).ToString();
                        //oSheet.Cells[current_rownum, 9].NumberFormat ="Percentage";
                        VAT_Rownum = current_rownum;
                        current_rownum++;

                        oSheet.Cells[current_rownum, 2] = "Thuế VAT";
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(226, 239, 218);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Cells[current_rownum, 9].Formula = "=I" + (current_rownum - 2).ToString() + "*H" + (current_rownum - 1).ToString();
                        oSheet.Cells[current_rownum, 7].Formula = "=G" + (current_rownum - 2).ToString() + "*H" + (current_rownum - 1).ToString();
                        current_rownum++;

                        oSheet.Cells[current_rownum, 2] = "Tổng giá trị thi công(bao gồm VAT)";
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(226, 239, 218);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Cells[current_rownum, 9].Formula = "=I" + (current_rownum - 1).ToString() + "+I" + (current_rownum - 3).ToString();
                        oSheet.Cells[current_rownum, 7].Formula = "=G" + (current_rownum - 1).ToString() + "+G" + (current_rownum - 3).ToString();
                        current_rownum++;

                        oSheet.Cells[current_rownum, 2] = "THANH TOÁN";
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Range["B" + current_rownum, "I" + current_rownum].Merge();
                        current_rownum++;

                        oSheet.Cells[current_rownum, 1] = "STT";
                        oSheet.Cells[current_rownum, 2] = "Nội dung";
                        oSheet.Cells[current_rownum, 4] = "%";
                        oSheet.Cells[current_rownum, 5] = "Giá trị";
                        oSheet.Cells[current_rownum, 4].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        oSheet.Range["B" + current_rownum, "C" + current_rownum].Merge();
                        oSheet.Range["B" + current_rownum, "C" + current_rownum].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        oSheet.Range["E" + current_rownum, "I" + current_rownum].Merge();
                        oSheet.Range["E" + current_rownum, "I" + current_rownum].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(189, 214, 238);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        current_rownum++;

                        oSheet.Cells[current_rownum, 1] = "1";
                        oSheet.Cells[current_rownum, 2] = "Tổng giá trị thi công (bao gồm VAT)";
                        oSheet.Cells[current_rownum, 4] = "";
                        oSheet.Range["E" + current_rownum].NumberFormat = "#,##0";
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=ROUND(G{0},0)", current_rownum - 3); //string.Format("=ROUND(SUM({0})*(1+{1}),0)", cell_sum_total, "H" + VAT_Rownum);
                        oSheet.Range["B" + current_rownum, "C" + current_rownum].Merge();
                        oSheet.Range["E" + current_rownum, "I" + current_rownum].Merge();
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        current_rownum++;

                        oSheet.Cells[current_rownum, 1] = "2";
                        oSheet.Cells[current_rownum, 2] = "Giá trị thực hiện đến kỳ này";

                        oSheet.Cells[current_rownum, 4] = "";
                        oSheet.Range["E" + current_rownum].NumberFormat = "#,##0";
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=ROUND(I{0},0)", current_rownum - 4);
                        oSheet.Range["B" + current_rownum, "C" + current_rownum].Merge();
                        oSheet.Range["E" + current_rownum, "I" + current_rownum].Merge();
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        current_rownum++;

                        float PTBH = 0, PTTU = 0, PTHU = 0, PTGL = 0;
                        //decimal GTTU = 0;
                        string TTTU = "",HTBH = "";
                        if (Additional_Info.Rows.Count > 0)
                        {
                            float.TryParse(Additional_Info.Rows[0]["PTBH"].ToString(), out PTBH);
                            float.TryParse(Additional_Info.Rows[0]["PTGL"].ToString(), out PTGL);
                            float.TryParse(Additional_Info.Rows[0]["PTTU"].ToString(), out PTTU);
                            float.TryParse(Additional_Info.Rows[0]["PTHU"].ToString(), out PTHU);
                            TTTU = Additional_Info.Rows[0]["TTTU"].ToString();
                            HTBH = Additional_Info.Rows[0]["HTBH"].ToString();
                            //decimal.TryParse(Additional_Info.Rows[0]["GTTU"].ToString(), out GTTU);
                        }


                        oSheet.Cells[current_rownum, 1] = "3";
                        oSheet.Cells[current_rownum, 2] = "Giá trị được thanh toán đến kỳ này";
                        oSheet.Range["D" + current_rownum].NumberFormat = "0.00%";
                        oSheet.Range["E" + current_rownum].NumberFormat = "#,##0";

                        if (BType == "3")
                        {
                            if (HTBH == "TM")
                            {
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 4]).Value2 = (1 - PTBH).ToString();
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=ROUND(E{0}*D{1},0)", current_rownum - 1, current_rownum);
                            }
                            else
                            {
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 4]).Value2 = (1).ToString();
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=ROUND(E{0}*D{1},0)", current_rownum - 1, current_rownum);
                            }
                        }
                        else
                        {
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 4]).Value2 = (1 - PTGL).ToString();
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=ROUND(E{0}*D{1},0)", current_rownum - 1, current_rownum);
                        }
                        oSheet.Range["B" + current_rownum, "C" + current_rownum].Merge();
                        oSheet.Range["E" + current_rownum, "I" + current_rownum].Merge();
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        current_rownum++;



                        oSheet.Cells[current_rownum, 1] = "4";
                        oSheet.Range["D" + current_rownum].NumberFormat = "0.00%";
                        oSheet.Range["E" + current_rownum].NumberFormat = "#,##0";
                        if (BType == "3")
                        {
                            
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 4]).Value2 = PTBH.ToString();
                            if (HTBH == "TM")
                            {
                                oSheet.Cells[current_rownum, 2] = "Giá trị giữ lại bảo hành";
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=ROUND(E{0}*D{1},0)", current_rownum - 2, current_rownum);
                            }
                            else
                            {
                                oSheet.Cells[current_rownum, 2] = "Giá trị giữ lại bảo hành (Chứng thư)";
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Value2 = 0;
                            }
                        }
                        else
                        {
                            oSheet.Cells[current_rownum, 2] = "Giá trị giữ lại kỳ này";
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 4]).Value2 = PTGL.ToString();
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=ROUND(E{0}*D{1},0)", current_rownum - 2, current_rownum);
                        }

                        //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 4]).Value2 =  PTBH.ToString();
                        //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=ROUND(E{0}*D{1},0)", current_rownum - 2, current_rownum);
                        oSheet.Range["B" + current_rownum, "C" + current_rownum].Merge();
                        oSheet.Range["E" + current_rownum, "I" + current_rownum].Merge();
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        current_rownum++;

                        oSheet.Cells[current_rownum, 1] = "5";
                        oSheet.Cells[current_rownum, 2] = "Tạm ứng";
                        oSheet.Range["D" + current_rownum].NumberFormat = "0.00%";
                        oSheet.Range["E" + current_rownum].NumberFormat = "#,##0";
                        oSheet.Cells[current_rownum, 4] = PTTU.ToString();
                        decimal pp_pl = 0, pp_ca = 0, TU_NEW = 0, pp_ca_no_VAT = 0, pp_tu_lastbill = 0, HU_NEW = 0, HU_NEW_LASTBILL = 0,PhiQL_LASTBILL=0;
                        if (Total_PP.Rows.Count == 1)
                        {
                            decimal.TryParse(Total_PP.Rows[0]["SUM_PL"].ToString(), out pp_pl);
                            decimal.TryParse(Total_PP.Rows[0]["SUM_CA"].ToString(), out pp_ca);
                            decimal.TryParse(Total_PP.Rows[0]["SUM_CA_NOVAT"].ToString(), out pp_ca_no_VAT);
                            decimal.TryParse(Total_PP.Rows[0]["TOTAL_TU"].ToString(), out TU_NEW);
                            decimal.TryParse(Total_PP.Rows[0]["TOTAL_TU_LASTBILL"].ToString(), out pp_tu_lastbill);
                            decimal.TryParse(Total_PP.Rows[0]["TOTAL_HU"].ToString(), out HU_NEW);
                            decimal.TryParse(Total_PP.Rows[0]["TOTAL_HU_LASTBILL"].ToString(), out HU_NEW_LASTBILL);
                            decimal.TryParse(Total_PP.Rows[0]["PhiQL"].ToString(), out PhiQL_LASTBILL);
                        }
                        if (BType == "1")
                        {
                            //if (period == 1)
                            dt_tmp.Rows[0]["U_GTTU"].ToString();
                            oSheet.Cells[current_rownum, 5].Value2 = dt_tmp.Rows[0]["U_GTTU"].ToString();
                            //else
                            //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=ROUND(SUM({0})*{1},0)", cell_sum_total, "D" + current_rownum);
                        }
                        else
                        {
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Value2 = TU_NEW;//string.Format("=ROUND(SUM({0})*{1},0)", cell_sum_total, "D" + current_rownum);
                        }
                        oSheet.Range["B" + current_rownum, "C" + current_rownum].Merge();
                        oSheet.Range["E" + current_rownum, "I" + current_rownum].Merge();
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        current_rownum++;

                        oSheet.Cells[current_rownum, 1] = "6";
                        oSheet.Cells[current_rownum, 2] = "Hoàn trả tạm ứng";
                        oSheet.Range["D" + current_rownum].NumberFormat = "0.00%";
                        oSheet.Range["E" + current_rownum].NumberFormat = @"_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)";
                        //oSheet.Cells[current_rownum, 4] = PTHU.ToString();
                        if (BType == "3")
                        {
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Value2 = -TU_NEW;
                        }
                        else
                        {
                            if (TU_NEW > 0)
                                //((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=-{0}*{1}", "I" + Tong_GT_RowNum, "D" + current_rownum);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Value2 = -HU_NEW;//string.Format("=-{0}*{1}", "I" + Tong_GT_RowNum, "D" + current_rownum);
                        }
                        oSheet.Range["B" + current_rownum, "C" + current_rownum].Merge();
                        oSheet.Range["E" + current_rownum, "I" + current_rownum].Merge();
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        current_rownum++;

                        oSheet.Cells[current_rownum, 1] = "7";
                        oSheet.Cells[current_rownum, 2] = "Tổng giá trị được thanh toán đến kỳ này";
                        oSheet.Range["E" + current_rownum].NumberFormat = "#,##0";
                        oSheet.Cells[current_rownum, 4] = "";
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=E{0}+E{1}+E{2}", current_rownum - 4, current_rownum - 2, current_rownum - 1);
                        oSheet.Range["B" + current_rownum, "C" + current_rownum].Merge();
                        oSheet.Range["E" + current_rownum, "I" + current_rownum].Merge();
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        current_rownum++;


                        //Lay so ky truoc
                        
                        oSheet.Cells[current_rownum, 1] = "8";
                        oSheet.Cells[current_rownum, 2] = "Tổng giá trị được thanh toán đến kỳ trước (bao gồm tạm ứng)";
                        oSheet.Cells[current_rownum, 4] = "";
                        oSheet.Range["E" + current_rownum].NumberFormat = "#,##0";
                        //if (TTTU == "01")
                        //{
                        //    if (period > 1)
                        //        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("={0} + {1} ", pp_ca, GTTU);
                        //    else
                        //        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("={0}", pp_ca);
                        //}
                        //else
                        //{
                        //    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("={0}", pp_ca);
                        //}
                        if (BType == "1")
                        {
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Value2 = 0;
                        }
                        else if (BType == "2")
                        {
                            if (pp_tu_lastbill > 0)
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("={0}*{1} + {2} - {3}", pp_ca + (pp_ca * (PhiQL_LASTBILL / 100)), (1 - PTGL).ToString(), pp_tu_lastbill, HU_NEW_LASTBILL); //pp_ca_no_VAT * (decimal)PTHU);
                            else
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("={0}*{1}", pp_ca + (pp_ca * (PhiQL_LASTBILL / 100)), (1 - PTGL).ToString());
                        }
                        else if (BType == "3")
                        {
                            if (pp_tu_lastbill > 0)
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("={0}*{1} + {2} - {3}", pp_ca + (pp_ca * (PhiQL_LASTBILL / 100)), (1 - PTGL).ToString(), pp_tu_lastbill, HU_NEW);//pp_ca_no_VAT * (decimal)PTHU);
                            else
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("={0}*{1}", pp_ca + (pp_ca * (PhiQL_LASTBILL / 100)), (1 - PTGL).ToString());
                        }
                        oSheet.Range["B" + current_rownum, "C" + current_rownum].Merge();
                        oSheet.Range["E" + current_rownum, "I" + current_rownum].Merge();
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        current_rownum++;

                        oSheet.Cells[current_rownum, 1] = "9";
                        oSheet.Cells[current_rownum, 2] = "Đề nghị thanh toán kỳ này (9) = (7) - (8)";
                        oSheet.Cells[current_rownum, 4] = "";
                        oSheet.Range["E" + current_rownum].NumberFormat = "#,##0";
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=E{0}-E{1}", current_rownum - 2, current_rownum - 1);
                        oSheet.Range["B" + current_rownum, "C" + current_rownum].Merge();
                        oSheet.Range["E" + current_rownum, "I" + current_rownum].Merge();
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(255, 229, 153);
                        current_rownum++;

                        decimal total = 0;
                        decimal.TryParse(oSheet.Cells[current_rownum - 1, 5].Value2.ToString(), out total);
                        oSheet.Cells[current_rownum, 5] = string.Format("(Bằng chữ: {0})", convert(total));
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Merge();
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(255, 229, 153);
                        current_rownum++;

                        //Border
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.get_Range("A8", "I" + (current_rownum - 1))).Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                        //Signature
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        oSheet.Cells[current_rownum, 2] = @"Ngày ..../..../....";
                        oSheet.Cells[current_rownum, 3] = @"Ngày ..../..../....";
                        oSheet.Cells[current_rownum, 6] = @"Ngày ..../..../....";
                        oSheet.Cells[current_rownum, 8] = @"Ngày ..../..../....";
                        if (BType=="3")
                        oSheet.Cells[current_rownum, 9] = @"Ngày ..../..../....";
                        current_rownum++;

                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        oSheet.Cells[current_rownum, 2] = "NTP/NCC/ĐTC";
                        oSheet.Cells[current_rownum, 3] = "Tổ chức thi công";
                        oSheet.Cells[current_rownum, 6] = "Quản lý thi công";
                        oSheet.Cells[current_rownum, 8] = "Chỉ huy trưởng";
                        if (BType == "3")
                        oSheet.Cells[current_rownum, 9] = "GĐ Dự án";
                        current_rownum++;

                        oApp.MessageBox("Export Excel Completed");
                    }
                }
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {

            }
        }

        //Delete
        private void Button6_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Grid0.Rows.SelectedRows.Count == 1)
            {
                //Get Data
                int DocEntry = 0;
                string docnum = Grid0.DataTable.GetValue("Document Number", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                int.TryParse(docnum, out DocEntry);
                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                SAPbobsCOM.CompanyService sCmp = null;
                sCmp = oCompany.GetCompanyService();
                oGeneralService = sCmp.GetGeneralService("KLTT");
                oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", DocEntry);
                oGeneralService.Delete(oGeneralParams);
                Load_Grid_Period(this.EditText1.Value.Trim(), ComboBox0.Selected.Value, ComboBox1.Selected.Value, ComboBox3.Selected.Value);
            }
        }

        private void Form_CloseBefore(SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (oCompany != null)
                    oCompany.Disconnect();
            }
            catch
            { }

        }
    }
}