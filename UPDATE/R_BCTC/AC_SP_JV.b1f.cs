using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;

namespace R_BCTC
{
    [FormAttribute("R_BCTC.AC_SP_JV", "AC_SP_JV.b1f")]
    class AC_SP_JV : UserFormBase
    {
        public AC_SP_JV()
        {
        }

        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        int DocEntry = 0, BatchNum = 0, Period = 0;
        string User_Create = "", Type = "", Dep_BpName = "", FProject = "";

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Grid Grid0;

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_1").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_Grid0").Specific));
            this.Grid0.PressedAfter += new SAPbouiCOM._IGridEvents_PressedAfterEventHandler(this.Grid0_PressedAfter);
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
            this.oApp = (SAPbouiCOM.Application)SAPbouiCOM.Framework.Application.SBO_Application;
            this.oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
            //Create Connection SQL
            SAPbobsCOM.Recordset or_RecoderSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            or_RecoderSet.DoQuery("Select * from [@ADDONCFG]");
            if (or_RecoderSet.RecordCount > 0)
            {
                string uid = or_RecoderSet.Fields.Item("Code").Value.ToString();
                string pwd = or_RecoderSet.Fields.Item("Name").Value.ToString();
                conn = new SqlConnection(string.Format("Data Source={0}; Initial Catalog={1}; User id={2}; Password={3};", oCompany.Server, oCompany.CompanyDB, uid, pwd));
            }
            else
            {
                oApp.MessageBox("Can't connect to DB !");
            }
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
            if (CheckExistUniqueID(oForm, "DT_JVList"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_JVList");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_JVList");
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
        
        private SAPbouiCOM.DataTable Convert_SAP_DataTable_Info(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_JVInfo"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_JVInfo");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_JVInfo");
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

        private System.Data.DataTable Get_MenuUID(string pReportName)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("GET_MENUUID_TV", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@ReportName", pReportName);
                cmd.CommandTimeout = 0;
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
        
        // Load Grid All JV
        private void Load_Grid_Period_All()
        {
            //LoadGrid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("SPJV_GetList_Current", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader dr = cmd.ExecuteReader();
                result.Load(dr);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();

                //Grid0.Columns.Item("POST_LVL").Visible = false;
                Grid0.Columns.Item("ProfitCode/BPCode").Visible = false;
                for (int i = 0; i < Grid0.Columns.Count; i++)
                {
                    Grid0.Columns.Item(i).Editable = false;
                }
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
        
        //Show all Button
        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Load_Grid_Period_All();
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }

        }

        //Press on Grid
        private void Grid0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Grid0.Rows.SelectedRows.Count == 1)
            {
                int.TryParse(Grid0.DataTable.GetValue("BatchNum", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out BatchNum);
                Period = 0;
                int.TryParse(Grid0.DataTable.GetValue("Period", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out Period);
                User_Create = Grid0.DataTable.GetValue("Username", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                Type = Grid0.DataTable.GetValue("Type", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                Dep_BpName = Grid0.DataTable.GetValue(@"Department/BPName", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                FProject = Grid0.DataTable.GetValue(@"Project", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            }

        }

        //Print Button
        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                DataTable rs = Get_MenuUID("Ủy nhiệm chi");
                if (rs.Rows.Count > 0)
                {
                    oApp.ActivateMenuItem(rs.Rows[0]["MenuUID"].ToString());
                    SAPbouiCOM.Form act_frm = oApp.Forms.ActiveForm;
                    ((SAPbouiCOM.EditText)act_frm.Items.Item("1000003").Specific).Value = BatchNum.ToString();
                    act_frm.Items.Item("1").Click();
                }
            }
            catch (Exception ex)
            {
                oApp.MessageBox("Error! " + ex.Message);
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

        }

    }
}
