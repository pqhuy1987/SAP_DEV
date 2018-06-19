using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using System.Data.SqlClient;
using System.Data;
using System.Globalization;
using System.Net.Mail;

namespace U_KLTT
{
    [FormAttribute("U_KLTT.Approve_Form", "Approve_Form.b1f")]
    class Approve_Form : UserFormBase
    {
        public Approve_Form()
        {
        }
        SAPbouiCOM.Application oApp = null;
        SAPbobsCOM.Company oCompany = null;
        SqlConnection conn = null;
        int DocEntry = 0;
        
        string SoHD = "";
        string NgayHD = "";
        string HD_Descript = "";
        //Email Server Config
        string Email_From = "";
        string Email_From_Name = "";
        string Host_Address = "";
        int Host_Port = 25;
        bool EnableSSL = false;
        string Uid = "";
        string Pwd = "";
        //End Email Config
        string FProject = "";
        string E_BpName = "";
        int E_Period = 0;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Grid Grid1;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.StaticText StaticText13;
        private SAPbouiCOM.StaticText StaticText14;
        private SAPbouiCOM.StaticText StaticText15;
        private SAPbouiCOM.StaticText StaticText18;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.EditText EditText9;
        private SAPbouiCOM.EditText EditText10;
        private SAPbouiCOM.EditText EditText11;
        private SAPbouiCOM.EditText EditText12;
        private SAPbouiCOM.EditText EditText13;
        private SAPbouiCOM.EditText EditText14;
        private SAPbouiCOM.EditText EditText16;
        private SAPbouiCOM.EditText EditText18;
        private SAPbouiCOM.EditText EditText20;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText16;
        private SAPbouiCOM.EditText EditText23;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.OptionBtn OptionBtn0;
        private SAPbouiCOM.OptionBtn OptionBtn1;
        private SAPbouiCOM.OptionBtn OptionBtn2;
        private SAPbouiCOM.OptionBtn OptionBtn3;
        
       
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("grd_lst").Specific));
            this.Grid0.PressedAfter += new SAPbouiCOM._IGridEvents_PressedAfterEventHandler(this.Grid0_PressedAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_4").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_5").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_13").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_14").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_15").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_16").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_17").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_18").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_19").Specific));
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_20").Specific));
            this.StaticText15 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_21").Specific));
            this.StaticText16 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_44").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("txt_gthd").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("txt_plhd").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("txt_hdtt").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("txt_gttc").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("txt_gtth").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("txt_pttt").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("txt_gttt").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("txt_ptgl").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("txt_gtgl").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("txt_pttu").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("txt_tu").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("txt_ptht").Specific));
            this.EditText14 = ((SAPbouiCOM.EditText)(this.GetItem("txt_hu").Specific));
            this.EditText16 = ((SAPbouiCOM.EditText)(this.GetItem("txt_tgt").Specific));
            this.EditText18 = ((SAPbouiCOM.EditText)(this.GetItem("txt_tgtt").Specific));
            this.EditText20 = ((SAPbouiCOM.EditText)(this.GetItem("txt_dntt").Specific));
            this.EditText23 = ((SAPbouiCOM.EditText)(this.GetItem("txt_note").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("bt_appr").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("bt_rej").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Grid1 = ((SAPbouiCOM.Grid)(this.GetItem("grd1").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("bt_cover").Specific));
            this.Button2.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button2_PressedAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("txt_dt").Specific));
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("bt_bill").Specific));
            this.Button3.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button3_PressedAfter);
            this.OptionBtn0 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_cur").Specific));
            this.OptionBtn0.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn0_PressedAfter);
            this.OptionBtn1 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_appr").Specific));
            this.OptionBtn1.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn1_PressedAfter);
            this.OptionBtn2 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_rej").Specific));
            this.OptionBtn2.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn2_PressedAfter);
            this.OptionBtn3 = ((SAPbouiCOM.OptionBtn)(this.GetItem("op_all").Specific));
            this.OptionBtn3.PressedAfter += new SAPbouiCOM._IOptionBtnEvents_PressedAfterEventHandler(this.OptionBtn3_PressedAfter);
            this.StaticText18 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_6").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("txt_link").Specific));
            this.EditText5.DoubleClickAfter += new SAPbouiCOM._IEditTextEvents_DoubleClickAfterEventHandler(this.EditText5_DoubleClickAfter);
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
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);
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
            
            //Load_Grid_Period();
            DataTable rs = Get_Email_Conf();
            if (rs.Rows.Count == 1)
            {
                Email_From = rs.Rows[0]["Email_From"].ToString();
                Email_From_Name = rs.Rows[0]["Email_From_Name"].ToString();
                Host_Address = rs.Rows[0]["Host_Address"].ToString();
                int.TryParse(rs.Rows[0]["Host_Port"].ToString(), out Host_Port);
                EnableSSL = false;
                if (rs.Rows[0]["EnableSSL"].ToString() == "0")
                    EnableSSL = false;
                else
                    EnableSSL = true;
                Uid = rs.Rows[0]["User"].ToString();
                Pwd = rs.Rows[0]["Pwd"].ToString();
            }
            OptionBtn1.GroupWith("op_cur");
            OptionBtn2.GroupWith("op_cur");
            OptionBtn3.GroupWith("op_cur");
            OptionBtn0.Selected = true;
        }

        private DataTable Get_Email_Conf()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("GET_EMAIL_CONF", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                if (conn.State == ConnectionState.Closed)
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

        private void Load_Grid_Period()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_GetList_Bill_Approve", conn);
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
                    if (c.TitleObject.Caption == "U_BPCode2")
                        c.Visible = false;
                    if (c.TitleObject.Caption == "POST_LVL")
                        c.Visible = false;
                    if (c.TitleObject.Caption == "Link")
                        c.Visible = false;
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

        private void Load_Grid_Period_Approved()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_GetList_Bill_Approved", conn);
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
                    if (c.TitleObject.Caption == "U_BPCode2")
                        c.Visible = false;
                    if (c.TitleObject.Caption == "POST_LVL")
                        c.Visible = false;
                    if (c.TitleObject.Caption == "Link")
                        c.Visible = false;
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

        private void Load_Grid_Period_Rejected()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_GetList_Bill_Rejected", conn);
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
                    if (c.TitleObject.Caption == "U_BPCode2")
                        c.Visible = false;
                    if (c.TitleObject.Caption == "POST_LVL")
                        c.Visible = false;
                    if (c.TitleObject.Caption == "Link")
                        c.Visible = false;
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

        private void Load_Grid_All()
        {
            //Load Grid
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_GetList_Bill_All", conn);
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
                    if (c.TitleObject.Caption == "U_BPCode2")
                        c.Visible = false;
                    if (c.TitleObject.Caption == "POST_LVL")
                        c.Visible = false;
                    if (c.TitleObject.Caption == "Link")
                        c.Visible = false;

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

        private void Load_Approve_Process()
        {
            //Load Grid
            DataTable result = Load_Approve_Process_KLTT(DocEntry);
            try
            {
                this.UIAPIRawForm.Freeze(true);
                Grid1.DataTable = Convert_SAP_DataTable_Approve_Process(result);
                Grid1.Columns.Item("U_Level").Visible = false;
                Grid1.Columns.Item("U_Position").Visible = false;
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                this.UIAPIRawForm.Freeze(false);
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
                { }

            }
            //Add row to DataTable
            foreach (System.Data.DataRow r in pDataTable.Rows)
            {
                oDT.Rows.Add();
                foreach (System.Data.DataColumn c in pDataTable.Columns)
                {
                    if (c.DataType.ToString() == "System.DateTime")
                        oDT.SetValue(c.ColumnName, oDT.Rows.Count - 1, r[c.ColumnName]);
                    else
                        oDT.SetValue(c.ColumnName, oDT.Rows.Count - 1, r[c.ColumnName].ToString());
                }
            }

            return oDT;
        }

        private SAPbouiCOM.DataTable Convert_SAP_DataTable_Approve_Process(System.Data.DataTable pDataTable)
        {
            SAPbouiCOM.DataTable oDT = null;
            SAPbouiCOM.Form oForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            if (CheckExistUniqueID(oForm, "DT_KLTTProcess"))
            {
                oDT = oForm.DataSources.DataTables.Item("DT_KLTTProcess");
                oDT.Clear();
            }
            else
            {
                oDT = oForm.DataSources.DataTables.Add("DT_KLTTProcess");
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

        private void Grid0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            //throw new System.NotImplementedException();
            //oApp.MessageBox(pVal.Row.ToString());
            
            DocEntry = 0;
            if (Grid0.Rows.SelectedRows.Count == 1)
            { 
                try
                {
                    //RESET
                    this.UIAPIRawForm.Freeze(true);
                    EditText0.Value = "0";
                    EditText1.Value = "0";
                    EditText2.Value = "0";
                    EditText3.Value = "0";
                    EditText4.Value = "0";
                    EditText6.Value = "0";
                    EditText7.Value = "0";
                    EditText8.Value = "0";
                    EditText9.Value = "0";
                    EditText10.Value = "0";
                    EditText11.Value = "0";
                    EditText12.Value = "0";
                    EditText13.Value = "0";
                    EditText14.Value = "0";
                    EditText16.Value = "0";
                    EditText18.Value = "0";
                    EditText20.Value = "0";
                    EditText23.Value = "";
                    EditText5.Value = "";
                    //find Selected Key Matrix
                    int.TryParse(Grid0.DataTable.GetValue("DocEntry", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out DocEntry);
                    FProject = Grid0.DataTable.GetValue("Project", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                    E_BpName = Grid0.DataTable.GetValue("BPName", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                    string BType = Grid0.DataTable.GetValue("Bill Type", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                    string BGroup = Grid0.DataTable.GetValue("BGroup", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                    string BPCode = Grid0.DataTable.GetValue("BPCode", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                    string PUType = Grid0.DataTable.GetValue("Purchase Type", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                    string SToDate = Grid0.DataTable.GetValue("To Date", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                    string U_BPCode2 = Grid0.DataTable.GetValue("U_BPCode2", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                    //string Link = Grid0.DataTable.GetValue("Link", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                    DateTime Todate = DateTime.Parse(SToDate);
                    int Period = 0;
                    int.TryParse(Grid0.DataTable.GetValue("Period", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out Period);
                    E_Period = Period;
                    if (BType == "Tạm ứng") BType = "1";
                    else if (BType == "Thanh toán") BType = "2";
                    else if (BType == "Quyết toán") BType = "3";
                    //Load Approve Process
                    Load_Approve_Process();
                    //Load Additional Info
                    DataTable SPP = Load_Data_KLTT(FProject, "SPP", Period - 1, BPCode, BGroup, PUType, DocEntry, Todate);
                    DataTable SPP_CURRENT = Load_Data_KLTT(FProject, "SPP", Period, BPCode, BGroup, PUType, DocEntry, Todate);
                    DataTable AI = Load_Data_KLTT(FProject, "AI", Period, BPCode, BGroup, PUType, DocEntry, Todate);
                    DataTable DUTRU_T = Load_Data_DUTRU(FProject, "AI", Period, BPCode, BGroup, PUType, DocEntry, Todate);

                    #region Thong tin chi tiet HD
                    decimal GTHD = 0, PLTT = 0, PLT = 0, DUTRU = 0;
                    float PTBH = 0, PTTU = 0, PTHU = 0, PTGL = 0;
                    string TTTU = "", HTBH = "";
                    if (AI.Rows.Count >= 1)
                    {
                        float.TryParse(AI.Rows[0]["PTBH"].ToString(), out PTBH);
                        float.TryParse(AI.Rows[0]["PTGL"].ToString(), out PTGL);
                        float.TryParse(AI.Rows[0]["PTTU"].ToString(), out PTTU);
                        float.TryParse(AI.Rows[0]["PTHU"].ToString(), out PTHU);
                        TTTU = AI.Rows[0]["TTTU"].ToString();
                        HTBH = AI.Rows[0]["HTBH"].ToString();
                        foreach (DataRow r in AI.Rows)
                        {
                            decimal tmp = 0;
                            decimal.TryParse(r["GTHD"].ToString(), out tmp);
                            if (r["Type"].ToString() == "HD")
                            {
                                GTHD += tmp;
                                SoHD = r["Number"].ToString();
                                NgayHD = r["StartDate"].ToString();
                                HD_Descript = r["Descript"].ToString();
                            }
                            else if (r["Type"].ToString() == "PLT")
                                PLT += tmp;
                            else if (r["Type"].ToString() == "PLTT")
                            {
                                PLTT += tmp;
                                SoHD = r["U_SHD"].ToString();
                                NgayHD = r["StartDate"].ToString();
                                HD_Descript = r["Descript"].ToString();
                            }
                        }
                    }
                    if (DUTRU_T.Rows.Count >= 1)
                    {
                        decimal.TryParse(DUTRU_T.Rows[0]["DUTRU"].ToString(), out DUTRU);
                    }
                    //GTHD
                    EditText0.Value = String.Format("{0:n0}", GTHD + PLTT);
                    //PL
                    EditText1.Value = String.Format("{0:n0}", PLT);
                    //GT HD sau dieu chinh
                    EditText2.Value = String.Format("{0:n0}", GTHD + PLTT + PLT);

                    if (BType == "3")
                    {
                        //Phan tram GT thanh toan den ky nay
                        EditText7.Value = (1 - PTBH).ToString();
                        //Phan tram giu lai
                        EditText9.Value = PTBH.ToString();
                    }
                    else
                    {
                        //Phan tram GT thanh toan den ky nay
                        EditText7.Value = (1 - PTGL).ToString();
                        //Phan tram giu lai
                        EditText9.Value = PTGL.ToString();
                    }
                  
                    //Phan tram tam ung
                    EditText11.Value = PTTU.ToString();
                    //Phan tram hoan tra
                    EditText13.Value = PTHU.ToString();
                    //Du tru
                    EditText3.Value = String.Format("{0:n0}", DUTRU);
                    #endregion

                    #region Thong tin bill
                    decimal GTTU = 0, GTHU = 0;
                    decimal pp_pl = 0, pp_ca = 0, pp_ca_no_VAT = 0, pp_tu_lastbill = 0, pp_hu_lastbill = 0;
                    decimal sum_ca_novat = 0;
                    decimal PhiQL = 0;
                    decimal PhiQL_LastBill = 0;
                    decimal ca = 0, pl_vat = 0;
                    if (SPP_CURRENT.Rows.Count == 1)
                    {
                        decimal.TryParse(SPP_CURRENT.Rows[0]["SUM_CA_NOVAT"].ToString(), out sum_ca_novat);
                        decimal.TryParse(SPP_CURRENT.Rows[0]["TOTAL_TU"].ToString(), out GTTU);
                        decimal.TryParse(SPP_CURRENT.Rows[0]["TOTAL_HU"].ToString(), out GTHU);
                        decimal.TryParse(SPP_CURRENT.Rows[0]["PhiQL"].ToString(), out PhiQL);
                        decimal.TryParse(SPP_CURRENT.Rows[0]["SUM_CA"].ToString(), out ca);
                        decimal.TryParse(SPP_CURRENT.Rows[0]["SUM_PL_VAT"].ToString(), out pl_vat);
                    }

                    if (SPP.Rows.Count == 1)
                    {
                        decimal.TryParse(SPP.Rows[0]["SUM_PL"].ToString(), out pp_pl);
                        decimal.TryParse(SPP.Rows[0]["SUM_CA"].ToString(), out pp_ca);
                        decimal.TryParse(SPP.Rows[0]["SUM_CA_NOVAT"].ToString(), out pp_ca_no_VAT);
                        //decimal.TryParse(SPP.Rows[0]["TOTAL_TU"].ToString(), out GTTU);
                        decimal.TryParse(SPP.Rows[0]["TOTAL_HU"].ToString(), out pp_hu_lastbill);
                        decimal.TryParse(SPP.Rows[0]["TOTAL_TU_LASTBILL"].ToString(), out pp_tu_lastbill);
                        decimal.TryParse(SPP.Rows[0]["PhiQL"].ToString(), out PhiQL_LastBill);
                    }
                    //Gia tri thuc hien den ky nay
                    EditText6.Value = BType == "1" ? "0" : String.Format("{0:n0}", ca * (1 + (PhiQL / 100)));
                    //Tong gia tri thi cong
                    EditText4.Value = BType == "1" ? "0" : String.Format("{0:n0}", pl_vat * (1 + (PhiQL / 100)));
                    //Tam ung
                    if (BType == "1")
                    {
                        decimal U_GTTU = 0;
                        string sql_cmd = string.Format("Select a.U_GTTU from [@KLTT] a where a.U_FIProject='{0}' and a.DocEntry = {1};", FProject, DocEntry);
                        try
                        {
                            SqlCommand cmd = new SqlCommand(sql_cmd, conn);
                            conn.Open();
                            decimal.TryParse(cmd.ExecuteScalar().ToString(), out U_GTTU);
                        }
                        catch
                        {

                        }
                        finally
                        {
                            conn.Close();
                        }
                        EditText12.Value = String.Format("{0:n0}", U_GTTU);
                    }
                    else
                    {
                        EditText12.Value = String.Format("{0:n0}", GTTU);
                    }
                    //Hoan tra TU
                    if (BType == "3")
                        EditText14.Value = String.Format("{0:n0}", GTTU);
                    else
                        EditText14.Value = String.Format("{0:n0}", GTHU);
                    //GT thanh toan den ky nay
                    if (BType == "3")
                        EditText8.Value = String.Format("{0:n0}", Math.Round(decimal.Parse(EditText6.Value), 0));
                    else
                        EditText8.Value = String.Format("{0:n0}", Math.Round((1 - (decimal)PTGL) * decimal.Parse(EditText6.Value), 0));
                    //GT thanh toan giu lai
                    if (BType == "3")
                    {
                        if (HTBH == "TM")
                        {
                            StaticText10.Caption = "4. GT giữ lại bảo hành";
                            EditText10.Value = String.Format("{0:n0}", Math.Round((decimal)PTBH * decimal.Parse(EditText6.Value), 0));
                        }
                        else
                        {
                            StaticText10.Caption = "4. GT giữ lại bảo hành (Chứng thư)";
                            EditText10.Value = "0";
                        }
                    }
                    else
                    {
                        StaticText10.Caption = "4. GT thanh toán giữ lại";
                        EditText10.Value = String.Format("{0:n0}", Math.Round((decimal)PTGL * decimal.Parse(EditText6.Value), 0));
                    }
                    //Tong GT duoc thanh toan den ky nay
                    EditText16.Value = String.Format("{0:n0}", decimal.Parse(EditText8.Value) + decimal.Parse(EditText12.Value) - decimal.Parse(EditText14.Value));
                    //Tong GT thanh toan den ky truoc
                    if (BType == "1")
                    {
                        EditText18.Value = "0";
                    }
                    else if (BType == "2")
                    {
                        if (pp_tu_lastbill > 0)
                            EditText18.Value = String.Format("{0:n0}", Math.Round((pp_ca * (1 - (decimal)PTGL)) + (pp_ca * (PhiQL_LastBill / 100)) + pp_tu_lastbill - pp_hu_lastbill, 0));
                        else
                            EditText18.Value = String.Format("{0:n0}", Math.Round((pp_ca * (1 - (decimal)PTGL)) + (pp_ca * (PhiQL_LastBill / 100)), 0));
                    }
                    else if (BType == "3")
                    {
                        if (pp_tu_lastbill > 0)
                            EditText18.Value = String.Format("{0:n0}", Math.Round((pp_ca * (1 - (decimal)PTGL)) + (pp_ca * (PhiQL_LastBill / 100)) + pp_tu_lastbill - pp_hu_lastbill));
                        else
                            EditText18.Value = String.Format("{0:n0}", Math.Round((pp_ca * (1 - (decimal)PTGL) + (pp_ca * (PhiQL_LastBill / 100)))));
                    }
                    //GT thanh toan ky nay
                    EditText20.Value = String.Format("{0:n0}", Math.Round(decimal.Parse(EditText16.Value) - decimal.Parse(EditText18.Value), 0));
                    //Link
                    //EditText5.Value = Link;
                    #endregion
                }
                catch
                {

                }
                finally
                {
                    this.UIAPIRawForm.Freeze(false);
                }
            }
        }

        System.Data.DataTable Load_Data_KLTT(string pFinancialProject, string pType, int pPeriod, string pBPCode, string pCGroup,string pPUType, int pDocEntry,DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                if (pType == "AI")
                {
                    cmd = new SqlCommand("KLTT_APPROVE_GET_ADDITIONALINFO", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                    cmd.Parameters.AddWithValue("@Period", pPeriod);
                    cmd.Parameters.AddWithValue("@BP_Code", pBPCode);
                    cmd.Parameters.AddWithValue("@CGroup", pCGroup);
                    cmd.Parameters.AddWithValue("@PUType", pPUType);
                    cmd.Parameters.AddWithValue("@ToDate", pToDate);
                }
                else if (pType == "SPP")
                {
                    cmd = new SqlCommand("KLTT_APPROVE_TOTAL", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                    cmd.Parameters.AddWithValue("@Period", pPeriod);
                    cmd.Parameters.AddWithValue("@BP_Code", pBPCode);
                    cmd.Parameters.AddWithValue("@BGroup", pCGroup);
                    cmd.Parameters.AddWithValue("@PUType", pPUType);
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

        System.Data.DataTable Load_Data_DUTRU(string pFinancialProject, string pType, int pPeriod, string pBPCode, string pCGroup, string pPUType, int pDocEntry, DateTime pToDate)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {

                cmd = new SqlCommand("KLTT_APPROVE_DUTRU", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                cmd.Parameters.AddWithValue("@BpCode", pBPCode);
                cmd.Parameters.AddWithValue("@CGroup", pCGroup);
                cmd.Parameters.AddWithValue("@PUType", pPUType);

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

        System.Data.DataTable Load_Approve_Process_KLTT(int pDocEntry)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_Approve_Process", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry", pDocEntry);
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

        private DataTable Get_lst_User_Next_LV()
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_Get_Lst_Usr_LV", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
                SqlDataReader rd = cmd.ExecuteReader();
                result.Load(rd);
            }
            catch (Exception ex)
            {
                oApp.MessageBox(ex.Message);
            }
            finally
            {
                cmd.Dispose();
            }
            return result;
        }

        private void Send_Email(string pToEMail, string pNameTo, string pBody, string pSubject)
        {
            MailMessage msg = new MailMessage();
            msg.To.Add(new MailAddress(pToEMail, pNameTo));
            msg.Priority = MailPriority.High;
            msg.From = new MailAddress(Email_From, Email_From_Name);
            msg.Subject = pSubject;
            msg.Body = pBody;
            msg.IsBodyHtml = true;

            SmtpClient client = new SmtpClient();
            client.UseDefaultCredentials = false;
            client.Credentials = new System.Net.NetworkCredential(Uid, Pwd);
            client.Port = Host_Port; // You can use Port 25 if 587 is blocked (mine is!)
            client.Host = Host_Address;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            //client.Timeout = 30;
            client.EnableSsl = EnableSSL;
            try
            {
                client.Send(msg);
            }
            catch (Exception ex)
            {
                oApp.SetStatusBarMessage("Send Email error: " + ex.Message);
            }
        }

        private void Send_Alert()
        {
            DataTable lst = Get_lst_User_Next_LV();
            if (lst.Rows.Count > 0)
            {
                SAPbobsCOM.Messages msg = null;
                msg = (SAPbobsCOM.Messages)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                msg.MessageText = string.Format("Bill số {0} của dự án {1} đang chờ duyệt.{2}Anh chị vui lòng vào xem.", DocEntry, FProject, Environment.NewLine);
                msg.Subject = "Yêu cầu phê duyệt Bill thanh toán số " + DocEntry.ToString();
                msg.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;
                for (int i = 0; i < lst.Rows.Count; i++)
                {
                    msg.Recipients.SetCurrentLine(i);
                    msg.Recipients.UserCode = lst.Rows[i]["USER_CODE"].ToString();
                    msg.Recipients.NameTo = lst.Rows[i]["NAME"].ToString();
                    msg.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                    if (lst.Rows[i]["EMAIL"].ToString() != "" && !string.IsNullOrEmpty(lst.Rows[i]["EMAIL"].ToString()))
                    {
                        string Body_Content = string.Format(@"Dear Anh/Chị,<br/><br/>"
                        + @"Có <b>đề nghị thanh toán</b> đang chờ bạn xử lý trên hệ thống SAP<br/><br/>"
                        + @"<b>Thông tin đề nghị thanh toán:</b><br/><br/>"
                        + @"P/B/BP/Dự án : <b>{0}</b><br/>"
                        + @"Kỳ thanh toán : <b>{1}</b><br/>"
                        + @"Số bill : <b>{2}</b><br/>"
                        + @"Đối tượng : <b>{3}</b><br/><br/>"
                        //+ "Trạng thái           : <b>{4}</b> đã duyệt<br/>"
                        + @"Đây là email được gửi tự động từ hệ thống SAP, vui lòng không trả lời lại email này. Xin cảm ơn.<br/><br/>"
                        + @"--------------<br/>"
                        + @"Trân trọng<br/>"
                        + @"SAP Business One",FProject,E_Period,DocEntry,E_BpName);
                        Send_Email(lst.Rows[i]["EMAIL"].ToString(), lst.Rows[i]["NAME"].ToString(), Body_Content, msg.Subject);
                    }
                    if (i < lst.Rows.Count - 1)
                    {
                        msg.Recipients.Add();
                    }
                }
                msg.Add();
            }
        }

        private void Send_Alert_Rejected()
        {
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(string.Format(@"Select USER_CODE, ISNULL(a.LastName,'') +' '+ ISNULL(a.MiddleName,'')+ ' '+ ISNULL(a.FirstName,'') as 'NAME',a.email--,a.empID,c.teamID,d.name "
                            + "from OHEM a inner join OUSR b on a.USERID = b.UserID "
                            + "where b.USERID in (Select UserSign from [@KLTT] where DocEntry={0})", DocEntry));
            if (oR_RecordSet.RecordCount > 0)
            {
                SAPbobsCOM.Messages msg = null;
                msg = (SAPbobsCOM.Messages)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oMessages);
                msg.MessageText = string.Format("Bill thanh toán số {0} đã bị từ chối", DocEntry);
                msg.Subject = "Bill thanh toán số " + DocEntry.ToString() + " bị từ chối";
                msg.Priority = SAPbobsCOM.BoMsgPriorities.pr_High;

                msg.Recipients.SetCurrentLine(0);
                msg.Recipients.UserCode = oR_RecordSet.Fields.Item("USER_CODE").Value.ToString();
                msg.Recipients.NameTo = oR_RecordSet.Fields.Item("NAME").Value.ToString();
                msg.Recipients.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;
                if (oR_RecordSet.Fields.Item("email").Value.ToString() != "" && !string.IsNullOrEmpty(oR_RecordSet.Fields.Item("email").Value.ToString()))
                    Send_Email(oR_RecordSet.Fields.Item("email").Value.ToString(), msg.Recipients.NameTo, msg.MessageText, msg.Subject);
                int kq = msg.Add();
                if (kq != 0)
                    oApp.SetStatusBarMessage("Send Email Result: " + oCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private bool Check_Ketoan_truong(string pUsrName, bool NT = false)
        {
            string sql = string.Format("Select position,dept from OHEM  where userID = (Select t.USERID from OUSR t where t.User_Code='{0}')",pUsrName);
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(sql);
            if (oR_RecordSet.RecordCount > 0)
            {
                string position = oR_RecordSet.Fields.Item("position").Value.ToString();
                string dept = oR_RecordSet.Fields.Item("dept").Value.ToString();
                if (NT == true)
                {
                    if (position == "3") return true;
                    else return false;
                }
                else
                {
                    if ((position == "1" && dept == "-2") || dept == "20" || dept == "21" || dept == "22") return true;
                    else return false;
                }
            }
            return false;
        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            int new_heigt = this.UIAPIRawForm.ClientHeight;
            int new_width = this.UIAPIRawForm.ClientWidth;

            Grid1.Item.Top = Grid0.Item.Top + Grid0.Item.Height + 20;
            Grid1.Item.Height = Button0.Item.Top - Grid1.Item.Top + 30;

        }

        System.Data.DataTable Load_Data_KLTT(string pFinancialProject, string pType, int pPeriod, string pBPCode, string pCGroup = "", int pDocEntry = 0, string pPUTYPE = "")
        {
            DataTable result = new DataTable();

            SqlCommand cmd = null;
            try
            {
                if (pType == "SPP")
                {
                    cmd = new SqlCommand("KLTT_TOTAL", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@FinancialProject", pFinancialProject);
                    cmd.Parameters.AddWithValue("@Period", pPeriod);
                    cmd.Parameters.AddWithValue("@BP_Code", pBPCode);
                    cmd.Parameters.AddWithValue("@BGroup", pCGroup);
                    cmd.Parameters.AddWithValue("@PUType", pPUTYPE);
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
                    rs += (rs == "" ? " " : " ") + ch[n[i]];// đọc số
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

        //Get MenuUID Crystal Report
        System.Data.DataTable Get_MenuUID(string pReportName)
        {
            DataTable result = new DataTable();
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("GET_MENUUID", conn);
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

        //Approve
        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SqlCommand cmd = null;
            string BPCode = Grid0.DataTable.GetValue("BPCode", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            bool NT = Check_NT(BPCode);
            try
            {
                cmd = new SqlCommand("KLTT_Approve_LV", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
                cmd.Parameters.AddWithValue("@Status", "1");
                cmd.Parameters.AddWithValue("@Comment", EditText23.Value.Trim());

                conn.Open();
                int rowupdate = cmd.ExecuteNonQuery();
                if (rowupdate >= 1)
                {
                    oApp.MessageBox("Phê duyệt thành công");
                    //Gui tin nhan
                    Send_Alert();
                    //Update Status neu la Ke toan truong hoac Nhan Tri
                    if (Check_Ketoan_truong(oCompany.UserName,NT))
                    {
                        SAPbobsCOM.GeneralService oGeneralService = null;
                        SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                        SAPbobsCOM.CompanyService sCmp = null;
                        sCmp = oCompany.GetCompanyService();
                        oGeneralService = sCmp.GetGeneralService("KLTT");
                        oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oGeneralParams.SetProperty("DocEntry", DocEntry);
                        oGeneralService.Close(oGeneralParams);
                    }
                }
                else
                {
                    oApp.MessageBox("Phê duyệt không thành công !");
                }
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
        }

        //Reject
        private void Button1_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SqlCommand cmd = null;
            try
            {
                cmd = new SqlCommand("KLTT_Approve_LV", conn);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@UserName", oCompany.UserName);
                cmd.Parameters.AddWithValue("@DocEntry", DocEntry);
                cmd.Parameters.AddWithValue("@Status", "2");
                cmd.Parameters.AddWithValue("@Comment", EditText23.Value.Trim());

                conn.Open();
                int rowupdate = cmd.ExecuteNonQuery();
                if (rowupdate >= 1)
                {
                    oApp.MessageBox("Reject thành công");
                    SAPbobsCOM.GeneralService oGeneralService = null;
                    SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                    SAPbobsCOM.CompanyService sCmp = null;
                    sCmp = oCompany.GetCompanyService();
                    oGeneralService = sCmp.GetGeneralService("KLTT");
                    oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oGeneralParams.SetProperty("DocEntry", DocEntry);
                    oGeneralService.Cancel(oGeneralParams);
                    //Send Alert
                    Send_Alert_Rejected();
                }
                else
                {
                    oApp.MessageBox("Phê duyệt không thành công!");
                }
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
        }

        //Print Cover
        private void Button2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            #region Excel
            //Microsoft.Office.Interop.Excel.Application oXL;
            //Microsoft.Office.Interop.Excel._Workbook oWB;
            //Microsoft.Office.Interop.Excel._Worksheet oSheet;
            //Microsoft.Office.Interop.Excel.Range oRng_Source;
            //Microsoft.Office.Interop.Excel.Range oRng_Dest;
            //object misvalue = System.Reflection.Missing.Value;
            //try
            //{
            //    if (Grid0.Rows.SelectedRows.Count == 1)
            //    {
            //        //Get DATA
            //        int.TryParse(Grid0.DataTable.GetValue("DocEntry", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out DocEntry);
            //        string FProject = Grid0.DataTable.GetValue("Project", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            //        string BType = Grid0.DataTable.GetValue("Bill Type", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            //        string BGroup = Grid0.DataTable.GetValue("BGroup", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            //        string BPCode = Grid0.DataTable.GetValue("BPCode", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            //        string BPName = Grid0.DataTable.GetValue("BPName", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            //        //string Name = Grid0.DataTable.GetValue("Creator Name", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            //        string PUType = Grid0.DataTable.GetValue("Purchase Type", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            //        string SToDate = Grid0.DataTable.GetValue("To Date", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            //        string U_BPCode2 = Grid0.DataTable.GetValue("U_BPCode2", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
            //        DateTime Todate = DateTime.Parse(SToDate);
            //        int Period = 0;
            //        int.TryParse(Grid0.DataTable.GetValue("Period", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out Period);
            //        if (BType == "Tạm ứng") BType = "1";
            //        else if (BType == "Thanh toán") BType = "2";
            //        else if (BType == "Quyết toán") BType = "3";

            //        //GTHD
            //        string GTHD = EditText0.Value;
            //        //PL
            //        string PL = EditText1.Value;
            //        //Tong Gia tri thi cong
            //        string GTTC = EditText4.Value;
            //        //Gia tri thuc hien den ky nay
            //        string GTTH = EditText6.Value;
            //        //Gia tri duoc thanh toan den ky nay
            //        string GTTT_KYNAY = EditText8.Value;
            //        //Tam ung
            //        string TU = EditText12.Value;
            //        //Hoan tam ung
            //        string HU = EditText14.Value;
            //        //Gia tri giu lai
            //        string GTGL = EditText10.Value;
            //        //GTTT ky truoc
            //        string GTTT_KYTRUOC = EditText18.Value;

            //        //Start Excel and get Application object.
            //        oXL = new Microsoft.Office.Interop.Excel.Application();
            //        oXL.Visible = true;
            //        //Open Template
            //        if (string.IsNullOrEmpty(U_BPCode2) || U_BPCode2 == "")
            //            oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\FDC_Bill_thanh_toan.xlsx");
            //        else
            //            oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\NT_Bill_thanh_toan.xlsx");
            //        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

            //        oSheet.Cells[2, 3] = "P/B/BP/DA: " + FProject;
            //        oSheet.Cells[3, 3] = "Ngày: " + DateTime.Now.ToString("dd/MM/yyyy");
            //        oSheet.Cells[4, 3] = "Số: " + Period;

            //        if (BPCode.Substring(0, 3) == "NCC")
            //        {
            //            if (BType == "3")
            //                oSheet.Cells[5, 3] = "THÔNG BÁO QUYẾT TOÁN NHÀ CUNG CẤP";
            //            else
            //                oSheet.Cells[5, 3] = "THÔNG BÁO THANH TOÁN NHÀ CUNG CẤP";
            //            oSheet.Cells[6, 1] = "Tên NCC: " + BPName;
            //            oSheet.Cells[7, 1] = "Số HĐ trên hệ thống: " + SoHD;
            //            oSheet.Cells[7, 6] = string.Format("Ngày: {0}", string.IsNullOrEmpty(NgayHD) ? "" : DateTime.Parse(NgayHD).ToString("dd/MM/yyyy"));
            //        }
            //        else if (BPCode.Substring(0, 3) == "NTP")
            //        {
            //            if (BType == "3")
            //                oSheet.Cells[5, 3] = "THÔNG BÁO QUYẾT TOÁN NHÀ THẦU PHỤ";
            //            else
            //                oSheet.Cells[5, 3] = "THÔNG BÁO THANH TOÁN NHÀ THẦU PHỤ";
            //            oSheet.Cells[6, 1] = "Tên NTP: " + BPName;
            //            oSheet.Cells[7, 1] = "Số HĐ trên hệ thống: " + SoHD;
            //            oSheet.Cells[7, 6] = string.Format("Ngày: {0}", string.IsNullOrEmpty(NgayHD) ? "" : DateTime.Parse(NgayHD).ToString("dd/MM/yyyy"));
            //        }
            //        else if (BPCode.Substring(0, 3) == "DTC")
            //        {
            //            if (BType == "3")
            //                oSheet.Cells[5, 3] = "THÔNG BÁO QUYẾT TOÁN ĐỘI THI CÔNG";
            //            else
            //                oSheet.Cells[5, 3] = "THÔNG BÁO THANH TOÁN ĐỘI THI CÔNG";
            //            oSheet.Cells[6, 1] = "Tên ĐTC: " + BPName;
            //            oSheet.Cells[7, 1] = "Số HĐ trên hệ thống: " + SoHD;
            //            oSheet.Cells[7, 6] = string.Format("Ngày: {0}", string.IsNullOrEmpty(NgayHD) ? "" : DateTime.Parse(NgayHD).ToString("dd/MM/yyyy"));
            //        }
            //        //Goi thau - HD Descript
            //        oSheet.Cells[8, 1] = "Gói thầu: " + HD_Descript;
            //        //GTHD
            //        oSheet.Cells[10, 6] = GTHD;
            //        //PL
            //        oSheet.Cells[11, 6] = PL;
            //        //Tong Gia tri thi cong
            //        oSheet.Cells[14, 6] = GTTC;
            //        //Gia tri thuc hien den ky nay
            //        oSheet.Cells[15, 6] = GTTH;
            //        //Gia tri duoc thanh toan den ky nay
            //        oSheet.Cells[16, 6] = GTTT_KYNAY;
            //        //Tam ung
            //        oSheet.Cells[17, 6] = TU;
            //        //Hoan tam ung
            //        oSheet.Cells[18, 6] = HU;
            //        //Gia tri giu lai
            //        if (BType == "3")
            //            oSheet.Cells[19, 2] = "Giá trị giữ lại bảo lãnh";
            //        else
            //            oSheet.Cells[19, 2] = "Giá trị giữ lại thanh toán";
            //        oSheet.Cells[19, 6] = GTGL;
            //        //GTTT ky truoc
            //        oSheet.Cells[21, 6] = GTTT_KYTRUOC;


            //        //oSheet.Cells[12, 7] = GTTT_KYNAY;
            //        //oSheet.Cells[13, 7] = GTTT_KYTRUOC;
            //        //Approve Process
            //        int current_row = 24;
            //        int current_col = 1;
            //        int block_count = 0;
            //        //Comment
            //        string Comment_Total = "Ghi chú:" + Environment.NewLine;
            //        for (int i = 0; i < Grid1.DataTable.Rows.Count; i++)
            //        {
            //            string Level = Grid1.DataTable.GetValue("U_Level", i).ToString();
            //            string Posistion = Grid1.DataTable.GetValue("U_Position", i).ToString();
            //            string Comment = Grid1.DataTable.GetValue("Comment", i).ToString();
            //            string ApprovedBy = Grid1.DataTable.GetValue("Approved by", i).ToString();
            //            string ApprovedOn = Grid1.DataTable.GetValue("Approved on", i).ToString();
            //            if (Level == "3")
            //            {
            //                if (Posistion == "5")
            //                {
            //                    if (!string.IsNullOrEmpty(Comment))
            //                        Comment_Total += string.Format("  CHT XD: {0}{1}", Comment, Environment.NewLine);
            //                }
            //                if (Posistion == "6")
            //                {
            //                    if (!string.IsNullOrEmpty(Comment))
            //                        Comment_Total += string.Format("  CHT ME: {0}{1}", Comment, Environment.NewLine);
            //                }
            //            }
            //            else if (Level == "1" && Posistion == "1")
            //            {
            //                if (!string.IsNullOrEmpty(Comment))
            //                    Comment_Total += string.Format("  CCM: {0}{1}", Comment, Environment.NewLine);
            //            }
            //            else if (Level == "6")
            //            {
            //                if (!string.IsNullOrEmpty(Comment))
            //                    Comment_Total += string.Format("  Giám đốc Dự án: {0}{1}", Comment, Environment.NewLine);
            //            }
            //            else if (Level == "-2" && Posistion == "1")
            //            {
            //                if (!string.IsNullOrEmpty(Comment))
            //                    Comment_Total += string.Format("  Kế toán - Tài chính: {0}{1}", Comment, Environment.NewLine);
            //            }
            //        }
            //        //Comment
            //        oSheet.Cells[23, 1] = Comment_Total;

            //        //Block
            //        block_count = 1;
            //        int block_row = 1;
            //        //current_row += 3;
            //        //current_col
            //        for (int i = 0; i < Grid1.DataTable.Rows.Count; i++)
            //        {
            //            string Level = Grid1.DataTable.GetValue("U_Level", i).ToString();
            //            string Posistion = Grid1.DataTable.GetValue("U_Position", i).ToString();
            //            string Comment = Grid1.DataTable.GetValue("Comment", i).ToString();
            //            string ApprovedBy = Grid1.DataTable.GetValue("Approved by", i).ToString();
            //            string ApprovedOn = Grid1.DataTable.GetValue("Approved on", i).ToString();
            //            string Status = Grid1.DataTable.GetValue("Status", i).ToString();
            //            if (Level == "3")
            //            {
            //                if (Posistion == "5")
            //                {
            //                    if (block_count > 4)
            //                    {
            //                        if (block_count > block_row * 4)
            //                        {
            //                            block_row++;
            //                            current_row += 8;
            //                        }
            //                        //Copy Block from Last Block
            //                        oRng_Source = oSheet.Range["A" + (current_row - 8) + ":B" + (current_row - 1)];
            //                        oRng_Source.Copy(oSheet.Cells[current_row, current_col]);
            //                    }
            //                    oSheet.Cells[current_row, current_col] = "Chỉ huy trưởng";
            //                    oSheet.Cells[current_row + 5, current_col] = Status;
            //                    oSheet.Cells[current_row + 6, current_col] = ApprovedBy;
            //                    oSheet.Cells[current_row + 7, current_col] = string.IsNullOrEmpty(ApprovedOn) ? "" : DateTime.ParseExact(ApprovedOn, "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy HH:mm:ss");
            //                    current_col += 2;
            //                    block_count++;

            //                }
            //                if (Posistion == "6")
            //                {
            //                    if (block_count > 4)
            //                    {
            //                        if (block_count > block_row * 4)
            //                        {
            //                            block_row++;
            //                            current_row += 8;
            //                        }
            //                        //Copy Block from Last Block
            //                        oRng_Source = oSheet.Range["A" + (current_row - 8) + ":B" + (current_row - 1)];
            //                        oRng_Source.Copy(oSheet.Cells[current_row, current_col]);

            //                    }
            //                    oSheet.Cells[current_row, current_col] = "Chỉ huy trưởng ME";
            //                    oSheet.Cells[current_row + 5, current_col] = Status;
            //                    oSheet.Cells[current_row + 6, current_col] = ApprovedBy;
            //                    oSheet.Cells[current_row + 7, current_col] = string.IsNullOrEmpty(ApprovedOn) ? "" : DateTime.ParseExact(ApprovedOn, "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy HH:mm:ss");
            //                    current_col += 2;
            //                    block_count++;

            //                }

            //            }
            //            else if (Level == "1" && Posistion == "1")
            //            {
            //                if (block_count > 4)
            //                {
            //                    if (block_count > block_row * 4)
            //                    {
            //                        block_row++;
            //                        current_row += 8;
            //                    }
            //                    //Copy Block from Last Block
            //                    oRng_Source = oSheet.Range["A" + (current_row - 8) + ":B" + (current_row - 1)];
            //                    oRng_Source.Copy(oSheet.Cells[current_row, current_col]);
            //                }
            //                oSheet.Cells[current_row, current_col] = "BKSCP";
            //                oSheet.Cells[current_row + 5, current_col] = Status;
            //                oSheet.Cells[current_row + 6, current_col] = ApprovedBy;
            //                oSheet.Cells[current_row + 7, current_col] = string.IsNullOrEmpty(ApprovedOn) ? "" : DateTime.ParseExact(ApprovedOn, "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy HH:mm:ss");
            //                current_col += 2;
            //                block_count++;

            //            }
            //            else if (Level == "6")
            //            {
            //                if (block_count > 4)
            //                {
            //                    if (block_count > block_row * 4)
            //                    {
            //                        block_row++;
            //                        current_row += 8;
            //                    }
            //                    //Copy Block from Last Block
            //                    oRng_Source = oSheet.Range["A" + (current_row - 8) + ":B" + (current_row - 1)];
            //                    oRng_Source.Copy(oSheet.Cells[current_row, current_col]);
            //                }
            //                oSheet.Cells[current_row, current_col] = "Giám đốc Dự Án";
            //                if (string.IsNullOrEmpty(U_BPCode2) || U_BPCode2 == "")
            //                    oSheet.Cells[current_row + 6, current_col] = ApprovedBy;
            //                oSheet.Cells[current_row + 5, current_col] = Status;
            //                oSheet.Cells[current_row + 7, current_col] = string.IsNullOrEmpty(ApprovedOn) ? "" : DateTime.ParseExact(ApprovedOn, "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy HH:mm:ss");
            //                current_col += 2;
            //                block_count++;

            //            }
            //            else if (Level == "-2" && Posistion == "1")
            //            {
            //                if (block_count > 4)
            //                {
            //                    if (block_count > block_row * 4)
            //                    {
            //                        block_row++;
            //                        current_row += 8;
            //                    }
            //                    //Copy Block from Last Block
            //                    oRng_Source = oSheet.Range["A" + (current_row - 8) + ":B" + (current_row - 1)];
            //                    oRng_Source.Copy(oSheet.Cells[current_row, current_col]);

            //                }
            //                oSheet.Cells[current_row, current_col] = "Kế toán trưởng";
            //                oSheet.Cells[current_row + 5, current_col] = Status;
            //                oSheet.Cells[current_row + 6, current_col] = ApprovedBy;
            //                oSheet.Cells[current_row + 7, current_col] = string.IsNullOrEmpty(ApprovedOn) ? "" : DateTime.ParseExact(ApprovedOn, "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy HH:mm:ss");
            //                current_col += 2;
            //                block_count++;
            //            }
            //            else if (Level == "16")
            //            {
            //                if (block_count > 4)
            //                {
            //                    if (block_count > block_row * 4)
            //                    {
            //                        block_row++;
            //                        current_row += 8;
            //                    }
            //                    //Copy Block from Last Block
            //                    oRng_Source = oSheet.Range["A" + (current_row - 8) + ":B" + (current_row - 1)];
            //                    oRng_Source.Copy(oSheet.Cells[current_row, current_col]);

            //                }
            //                oSheet.Cells[current_row, current_col] = "Kế toán trưởng";
            //                //oSheet.Cells[current_row + 6, current_col] = ApprovedBy;
            //                oSheet.Cells[current_row + 5, current_col] = Status;
            //                oSheet.Cells[current_row + 7, current_col] = string.IsNullOrEmpty(ApprovedOn) ? "" : DateTime.ParseExact(ApprovedOn, "dd MMM yyyy HH:mm:ss:fff", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy HH:mm:ss");
            //                current_col += 2;
            //                block_count++;
            //            }

            //            if (current_col > 7) current_col = 1;
            //        }
                    

            //        oApp.MessageBox("Export Excel Completed");
            //    }
            //}
            //catch (Exception ex)
            //{ 
            //    oApp.MessageBox(ex.Message); 
            //}
        #endregion
            #region Crystal Report
            try
            {
                if (Grid0.Rows.SelectedRows.Count == 1)
                {
                    int.TryParse(Grid0.DataTable.GetValue("DocEntry", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString(), out DocEntry);
                    string U_BPCode2 = Grid0.DataTable.GetValue("U_BPCode2", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();

                    DataTable rs = new DataTable();
                    if (string.IsNullOrEmpty(U_BPCode2))
                        rs = Get_MenuUID("Cover_FDC");
                    else
                        rs = Get_MenuUID("Cover_NT");
                    if (rs.Rows.Count > 0)
                    {
                        oApp.ActivateMenuItem(rs.Rows[0]["MenuUID"].ToString());
                        SAPbouiCOM.Form act_frm = oApp.Forms.ActiveForm;
                        ((SAPbouiCOM.EditText)act_frm.Items.Item("1000003").Specific).Value = DocEntry.ToString();
                        act_frm.Items.Item("1").Click();
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
            #endregion
        }

        //Print Bill
        private void Button3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            try
            {
                if (Grid0.Rows.SelectedRows.Count == 1)
                {
                    //Get Data
                    int DocEntry = 0;
                    string docnum = Grid0.DataTable.GetValue("DocEntry", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();
                    int.TryParse(docnum, out DocEntry);
                    string projectKey = Grid0.DataTable.GetValue("Project", Grid0.Rows.SelectedRows.Item(0, SAPbouiCOM.BoOrderType.ot_RowOrder)).ToString();

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
                        decimal PhiQL = 0;
                        decimal.TryParse(dt_tmp.Rows[0]["U_PTQuanly"].ToString(), out PhiQL);

                        DataTable A = Load_Data_KLTT(financialproject, "A", period, bp, BGroup, DocEntry);
                        DataTable B = Load_Data_KLTT(financialproject, "B", period, bp, BGroup, DocEntry);
                        DataTable C = Load_Data_KLTT(financialproject, "C", period, bp, BGroup, DocEntry);
                        DataTable D = Load_Data_KLTT(financialproject, "D", period, bp, BGroup, DocEntry);
                        DataTable E = Load_Data_KLTT(financialproject, "E", period, bp, BGroup, DocEntry);
                        DataTable F = Load_Data_KLTT(financialproject, "F", period, bp, BGroup, DocEntry);
                        DataTable G = Load_Data_KLTT(financialproject, "G", period, bp, BGroup, DocEntry);
                        DataTable H = Load_Data_KLTT(financialproject, "H", period, bp, BGroup, DocEntry);
                        DataTable K = Load_Data_KLTT(financialproject, "K", period, bp, BGroup, DocEntry);
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
                        //oWB = oXL.Workbooks.Open(System.Windows.Forms.Application.StartupPath + @"\Template_KLTT.xlsx");
                        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                        //Fill Header

                        //LOGO
                        if (QLNT == "NTP00599")
                        {
                            oSheet.Cells[1, 1] = "NHÂN TRÍ";
                        }
                        else if (QLNT == "NTP00601")
                        {
                            oSheet.Cells[1, 1] = "NHÂN TIẾN";
                        }
                        else if (QLNT == "NTP00602")
                        {
                            oSheet.Cells[1, 1] = "NHÂN TÍN";
                        }
                        //Tilte
                        if (BType == "3")
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
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 5]).Formula = "=SUM(E" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":E" + (current_rownum - 1).ToString() + ")";
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 7]).Formula = "=SUM(G" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":G" + (current_rownum - 1).ToString() + ")";
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 9]).Formula = "=SUM(I" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":I" + (current_rownum - 1).ToString() + ")";
                                                                    }

                                                                    #endregion

                                                                    //Total Level 4
                                                                    if (Gr4_element.Count > 0)
                                                                    {
                                                                        string cell_sum_tt = "";
                                                                        string cell_sum_gtth = "";
                                                                        string cell_sum_gtth2 = "";
                                                                        int temp = 1;
                                                                        foreach (int t in Gr4_element)
                                                                        {
                                                                            if (temp < Gr4_element.Count)
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
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                                                            string cell_sum_tt = "";
                                                            string cell_sum_gtth = "";
                                                            string cell_sum_gtth2 = "";
                                                            int temp = 1;
                                                            foreach (int t in Gr3_element)
                                                            {
                                                                if (temp < Gr3_element.Count)
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
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 5]).Formula = "=SUM(E" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":E" + (current_rownum - 1).ToString() + ")";
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 7]).Formula = "=SUM(G" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":G" + (current_rownum - 1).ToString() + ")";
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 9]).Formula = "=SUM(I" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":I" + (current_rownum - 1).ToString() + ")";
                                                                    }

                                                                    #endregion

                                                                    //Total Level 4
                                                                    if (Gr4_element.Count > 0)
                                                                    {
                                                                        string cell_sum_tt = "";
                                                                        string cell_sum_gtth = "";
                                                                        string cell_sum_gtth2 = "";
                                                                        int temp = 1;
                                                                        foreach (int t in Gr4_element)
                                                                        {
                                                                            if (temp < Gr4_element.Count)
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
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
                                                                    }
                                                                }

                                                            }
                                                        }
                                                        #endregion

                                                        //Total level 3
                                                        if (Gr3_element.Count > 0)
                                                        {
                                                            string cell_sum_tt = "";
                                                            string cell_sum_gtth = "";
                                                            string cell_sum_gtth2 = "";
                                                            int temp = 1;
                                                            foreach (int t in Gr3_element)
                                                            {
                                                                if (temp < Gr3_element.Count)
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
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUM(E{0}:E{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUM(G{0}:G{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUM(I{0}:I{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
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
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUM(E{0}:E{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUM(G{0}:G{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUM(I{0}:I{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
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
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUM(E{0}:E{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUM(G{0}:G{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUM(I{0}:I{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
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
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUM(E{0}:E{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUM(G{0}:G{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUM(I{0}:I{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
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
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUM(E{0}:E{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUM(G{0}:G{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUM(I{0}:I{1})", (Section_RowNum[(Section_RowNum.Count - 1)] + 1), current_rownum - 1);
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
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 5]).Formula = "=SUM(E" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":E" + (current_rownum - 1).ToString() + ")";
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 7]).Formula = "=SUM(G" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":G" + (current_rownum - 1).ToString() + ")";
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum5[(Group_No_RowNum5.Count - 1)], 9]).Formula = "=SUM(I" + (Group_No_RowNum5[(Group_No_RowNum5.Count - 1)] + 1) + ":I" + (current_rownum - 1).ToString() + ")";
                                                                    }

                                                                    #endregion

                                                                    //Total Level 4
                                                                    if (Gr4_element.Count > 0)
                                                                    {
                                                                        string cell_sum_tt = "";
                                                                        string cell_sum_gtth = "";
                                                                        string cell_sum_gtth2 = "";
                                                                        int temp = 1;
                                                                        foreach (int t in Gr4_element)
                                                                        {
                                                                            if (temp < Gr4_element.Count)
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
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                                                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum4[(Group_No_RowNum4.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
                                                                    }
                                                                }

                                                            }
                                                        }
                                                        #endregion

                                                        //Total level 3
                                                        if (Gr3_element.Count > 0)
                                                        {
                                                            string cell_sum_tt = "";
                                                            string cell_sum_gtth = "";
                                                            string cell_sum_gtth2 = "";
                                                            int temp = 1;
                                                            foreach (int t in Gr3_element)
                                                            {
                                                                if (temp < Gr3_element.Count)
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
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum3[(Group_No_RowNum3.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum2[(Group_No_RowNum2.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                    ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Group_No_RowNum[(Group_No_RowNum.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 5]).Formula = string.Format("=SUM({0})", cell_sum_tt);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 7]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                                ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[Section_RowNum[(Section_RowNum.Count - 1)], 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth2);
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
                            Section_RowNum.Add(current_rownum);
                            current_rownum++;
                        }

                        //TOTAL
                        int Tong_GT_RowNum = current_rownum;
                        oSheet.Cells[current_rownum, 2] = "TỔNG GIÁ TRỊ (Chưa VAT)";
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Interior.Color = System.Drawing.Color.FromArgb(226, 239, 218);
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        string cell_sum_total = "";
                        if (Section_RowNum.Count > 0)
                        {
                            string cell_sum_gtth = "";
                            int temp = 1;
                            foreach (int t in Section_RowNum)
                            {
                                if (temp < Section_RowNum.Count)
                                {
                                    cell_sum_gtth += "I" + t + ",";
                                    cell_sum_total += "G" + t + ",";
                                    temp++;
                                }
                                else
                                {
                                    cell_sum_gtth += "I" + t;
                                    cell_sum_total += "G" + t;
                                }
                            }
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 9]).Formula = string.Format("=SUM({0})", cell_sum_gtth);
                            ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 7]).Formula = string.Format("=SUM({0})", cell_sum_total);
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
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Formula = string.Format("=ROUND(SUM({0})*(1+{1}),0)", cell_sum_total, "H" + VAT_Rownum);
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
                        string TTTU = "", HTBH = "";
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
                        decimal pp_pl = 0, pp_ca = 0, TU_NEW = 0, pp_ca_no_VAT = 0, pp_tu_lastbill = 0, HU_NEW = 0, HU_NEW_LASTBILL = 0;
                        if (Total_PP.Rows.Count == 1)
                        {
                            decimal.TryParse(Total_PP.Rows[0]["SUM_PL"].ToString(), out pp_pl);
                            decimal.TryParse(Total_PP.Rows[0]["SUM_CA"].ToString(), out pp_ca);
                            decimal.TryParse(Total_PP.Rows[0]["SUM_CA_NOVAT"].ToString(), out pp_ca_no_VAT);
                            decimal.TryParse(Total_PP.Rows[0]["TOTAL_TU"].ToString(), out TU_NEW);
                            decimal.TryParse(Total_PP.Rows[0]["TOTAL_TU_LASTBILL"].ToString(), out pp_tu_lastbill);
                            decimal.TryParse(Total_PP.Rows[0]["TOTAL_HU"].ToString(), out HU_NEW);
                            decimal.TryParse(Total_PP.Rows[0]["TOTAL_HU_LASTBILL"].ToString(), out HU_NEW_LASTBILL);
                        }
                        oSheet.Cells[current_rownum, 5].Value2 = EditText12.Value;
                        oSheet.Range["B" + current_rownum, "C" + current_rownum].Merge();
                        oSheet.Range["E" + current_rownum, "I" + current_rownum].Merge();
                        oSheet.Range["A" + current_rownum, "I" + current_rownum].Font.Bold = true;
                        current_rownum++;

                        oSheet.Cells[current_rownum, 1] = "6";
                        oSheet.Cells[current_rownum, 2] = "Hoàn trả tạm ứng";
                        oSheet.Range["D" + current_rownum].NumberFormat = "0.00%";
                        oSheet.Range["E" + current_rownum].NumberFormat = @"_(* #,##0_);_(* (#,##0);_(* "" - ""??_);_(@_)";
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Value2 = "-" + EditText14.Value;

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
                        ((Microsoft.Office.Interop.Excel.Range)oSheet.Cells[current_rownum, 5]).Value2 = EditText18.Value;

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
                        if (BType == "3")
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
                Load_Grid_Period_Approved();
                Button0.Item.Enabled = false;
                Button1.Item.Enabled = false;
            }
        }

        //Show Rejected
        private void OptionBtn2_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn2.Selected)
            {
                Load_Grid_Period_Rejected();
                Button0.Item.Enabled = false;
                Button1.Item.Enabled = false;
            }
        }

        //Show All
        private void OptionBtn3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (OptionBtn3.Selected)
            {
                Load_Grid_All();
                Button0.Item.Enabled = false;
                Button1.Item.Enabled = false;
            }
        }

        //Double click vao Link
        private void EditText5_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (!string.IsNullOrEmpty(EditText5.Value))
                    System.Diagnostics.Process.Start(EditText5.Value);
            }
            catch
            { }
        }

        private bool Check_NT(string pBpCode)
        {
            SAPbobsCOM.Recordset oR_RecordSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oR_RecordSet.DoQuery(string.Format("Select GroupCode from OCRD where CardCode='{0}'", pBpCode));
            if (oR_RecordSet.RecordCount > 0)
            {
                string BPGCode = oR_RecordSet.Fields.Item("GroupCode").Value.ToString();
                if (BPGCode == "112") return true;
                else return false;
            }
            return false;
        }

    }
}
