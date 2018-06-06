using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;

namespace PurchaseBlanketAgreement
{
    [FormAttribute("PurchaseBlanketAgreement.frm_Approvelst", "frm_Approvelst.b1f")]
    class frm_Approvelst : UserFormBase
    {
        public frm_Approvelst()
        {
        }
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        int BlanketNo = 0;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        /// 
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Grid Grid1;
        private SAPbouiCOM.OptionBtn OptionBtn0;
        private SAPbouiCOM.OptionBtn OptionBtn1;
        private SAPbouiCOM.OptionBtn OptionBtn2;
        private SAPbouiCOM.OptionBtn OptionBtn3;
        private SAPbouiCOM.Button Button0;
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("grd0").Specific));
            this.Grid0.PressedAfter += new SAPbouiCOM._IGridEvents_PressedAfterEventHandler(this.Grid0_PressedAfter);
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Grid1 = ((SAPbouiCOM.Grid)(this.GetItem("grd1").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_2").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.OptionBtn0 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_cur").Specific));
            this.OptionBtn0.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn0_PressedAfter);
            this.OptionBtn1 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_appr").Specific));
            this.OptionBtn1.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn1_PressedAfter);
            this.OptionBtn2 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_rej").Specific));
            this.OptionBtn2.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn2_PressedAfter);
            this.OptionBtn3 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_sa").Specific));
            this.OptionBtn3.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn3_PressedAfter);
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
            //Load_Grid_HDList();
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);
            OptionBtn1.GroupWith("op_cur");
            OptionBtn2.GroupWith("op_cur");
            OptionBtn3.GroupWith("op_cur");
            OptionBtn0.Selected = true;
        }
        
        //Show Current
        private void Load_Grid_HDList()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BL_Get_List", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();
                //Disable all grid column
                foreach (SAPbouiCOM.GridColumn c in Grid0.Columns)
                {
                    c.Editable = false;
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

        //Show Approved
        private void Load_Grid_HDList_Approved()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BL_Get_List_Approved", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();
                //Disable all grid column
                foreach (SAPbouiCOM.GridColumn c in Grid0.Columns)
                {
                    c.Editable = false;
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

        //Show Rejected
        private void Load_Grid_HDList_Rejected()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BL_Get_List_Rejected", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();
                //Disable all grid column
                foreach (SAPbouiCOM.GridColumn c in Grid0.Columns)
                {
                    c.Editable = false;
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

        //Show All
        private void Load_Grid_HDList_All()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BL_Get_List_All", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid0.DataTable = Convert_SAP_DataTable(result);

                Grid0.AutoResizeColumns();
                //Disable all grid column
                foreach (SAPbouiCOM.GridColumn c in Grid0.Columns)
                {
                    c.Editable = false;
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

        private void Load_Grid_Approve_Process()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("BL_Get_Aprrove_Process", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@BlanketNo", BlanketNo);
                conn.Open();
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
                this.UIAPIRawForm.Freeze(true);
                Grid1.DataTable = Convert_SAP_DataTable_Approve_Process(result);
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
            if (CheckExistUniqueID(oForm, "DT_HDList"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_HDList");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_HDList");
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
        
        private SAPbouiCOM.DataTable Convert_SAP_DataTable_Approve_Process(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_AP_Process"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_AP_Process");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_AP_Process");
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

        private void Grid0_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Grid0.Rows.SelectedRows.Count > 0)
            {
                oApp.ActivateMenuItem("39722");
                SAPbouiCOM.Form act_form = oApp.Forms.ActiveForm;
                //1250000004
                ((SAPbouiCOM.EditText)act_form.Items.Item("1250000004").Specific).Value = Grid0.DataTable.GetValue("Agreement No", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                act_form.Items.Item("1250000001").Click();
            }

        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            int new_heigt = this.UIAPIRawForm.ClientHeight;
            int new_width = this.UIAPIRawForm.ClientWidth;

            //Application.SBO_Application.MessageBox(this.UIAPIRawForm.Width.ToString() + "x" + this.UIAPIRawForm.Height.ToString());
            //Resize List JV
            Grid0.Item.Height = new_heigt/2 - 50;
            //Resize List Apprv
            Grid1.Item.Top = new_heigt / 2 -20;
            Grid1.Item.Height = Button0.Item.Top - Grid1.Item.Top -5;

        }

        private void Grid0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (Grid0.Rows.SelectedRows.Count > 0)
            {
                int.TryParse(Grid0.DataTable.GetValue("Agreement No", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out BlanketNo);
                Load_Grid_Approve_Process();
            }
        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (BlanketNo > 0)
            {
                frm_Approve frm = new frm_Approve(BlanketNo.ToString() , "S", oApp.Forms.ActiveForm);
                frm.Show();
            }

        }

        private void Form_CloseBefore(SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //throw new System.NotImplementedException();
            try
            {
                oCompany.Disconnect();
            }
            catch
            { }

        }

        //Show Current
        private void OptionBtn0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn0.Selected)
            {
                Load_Grid_HDList();
                Button0.Item.Enabled = true;
            }
        }

        //Show Approved
        private void OptionBtn1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn1.Selected)
            {
                Load_Grid_HDList_Approved();
                Button0.Item.Enabled = false;
            }

        }

        //Show Rejected
        private void OptionBtn2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn2.Selected)
            {
                Load_Grid_HDList_Rejected();
                Button0.Item.Enabled = false;
            }

        }

        //Show All
        private void OptionBtn3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn3.Selected)
            {
                Load_Grid_HDList_All();
                Button0.Item.Enabled = false;
            }
        }
    }
}
